using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;

namespace MsgReader.Mime.Decode;

/// <summary>
///     Utility class for dealing with encoded word strings<br />
///     <br />
///     EncodedWord encoded strings are only in ASCII, but can embed information
///     about characters in other character sets.<br />
///     <br />
///     It is done by specifying the character set, an encoding that maps from ASCII to
///     the correct bytes and the actual encoded string.<br />
///     <br />
///     It is specified in a format that is best summarized by a BNF:<br />
///     <c>"=?" character_set "?" encoding "?" encoded-text "?="</c><br />
/// </summary>
/// <example>
///     <c>=?ISO-8859-1?Q?=2D?=</c>
///     Here <c>ISO-8859-1</c> is the character set.<br />
///     <c>Q</c> is the encoding method (quoted-printable). <c>B</c> is also supported (Base 64).<br />
///     The encoded text is the <c>=2D</c> part which is decoded to a space.
/// </example>
internal static class EncodedWord
{
    #region Decode
    /// <summary>
    ///     Decode text that is encoded with the <see cref="EncodedWord" /> encoding.<br />
    ///     <br />
    ///     This method will decode any encoded-word found in the string.<br />
    ///     All parts which is not encoded will not be touched.<br />
    ///     <br />
    ///     From <a href="http://tools.ietf.org/html/rfc2047">RFC 2047</a>:<br />
    ///     <code>
    ///  Generally, an "encoded-word" is a sequence of printable ASCII
    ///  characters that begins with "=?", ends with "?=", and has two "?"s in
    ///  between.  It specifies a character set and an encoding method, and
    ///  also includes the original text encoded as graphic ASCII characters,
    ///  according to the rules for that encoding method.
    ///  </code>
    ///     Example:<br />
    ///     <c>=?ISO-8859-1?q?this=20is=20some=20text?= other text here</c>
    /// </summary>
    /// <remarks>
    ///     See <a href="http://tools.ietf.org/html/rfc2047#section-2">RFC 2047 section 2</a> "Syntax of encoded-words"
    ///     for more details
    /// </remarks>
    /// <param name="encodedWords">Source text. May be content which is not encoded.</param>
    /// <returns>Decoded text</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="encodedWords" /> is <see langword="null" /></exception>
    public static string Decode(string encodedWords)
    {
        if (encodedWords == null)
            throw new ArgumentNullException(nameof(encodedWords));

        // Notice that RFC2231 redefines the BNF to
        // encoded-word := "=?" charset ["*" language] "?" encoded-text "?="
        // but no usage of this BNF have been spotted yet. It is here to
        // ease debugging if such a case is discovered.

        // This is the regex that should fit the BNF
        // RFC Says that NO WHITESPACE is allowed in this encoding, but there are examples
        // where whitespace is there, and therefore this regex allows for such.
		const string encodedWordRegex = @"=\?(?<Charset>[\w]+[^\r\n\t\f\v\?]+?)\?(?<Encoding>\w)\?(?<Content>.*?)\?\=";
        // \w	Matches any word character including underscore. Equivalent to "[A-Za-z0-9_]".
        // \S	Matches any nonwhite space character. Equivalent to "[^ \f\n\r\t\v]".
        // +?   non-greedy equivalent to +
        // (?<NAME>REGEX) is a named group with name NAME and regular expression REGEX

        // Any amount of linear-space-white between 'encoded-word's,
        // even if it includes a CRLF followed by one or more SPACEs,
        // is ignored for the purposes of display.
        // http://tools.ietf.org/html/rfc2047#page-12
        // Define a regular expression that captures two encoded words with some whitespace between them
        const string replaceRegex = @"(?<first>" + encodedWordRegex + @")\s+(?<second>" + encodedWordRegex + ")";

        // Then, find an occurrence of such an expression, but remove the whitespace in between when found
        // Need to be done twice for encodings such as "=?UTF-8?Q?a?= =?UTF-8?Q?b?= =?UTF-8?Q?c?="
        // to be replaced correctly
        encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}");
        encodedWords = Regex.Replace(encodedWords, replaceRegex, "${first}${second}");

        // var decodedWords = encodedWords;
        var decodedWords = string.Empty;

        var matches = Regex.Matches(encodedWords, $"({encodedWordRegex})|(?<Unencoded>[^=]+)")
            .Cast<Match>()
            .Where(static m => m.Success)
            .Select(
                static m => new {
                    m.Value,
                    Content = m.Groups["Content"].Success ? m.Groups["Content"].Value : m.Groups["Unencoded"].Value,
                    Encoding = m.Groups["Encoding"].Value,
                    Charset = m.Groups["Charset"].Value,
                    IsEncoded = m.Groups["Content"].Success
                }
            )
            .ToList();

        if (matches.Count == 0)
            return encodedWords;
        
        for (var i = 0; i < matches.Count; i++)
        {
            var tempValue = matches[i].Value;
            var tempContent = matches[i].Content;
            var tempEncoding = matches[i].Encoding;
            var tempCharset = matches[i].Charset;
            var isEncoded = matches[i].IsEncoded;

            // Store decoded text here when done
            string decodedText;

            if (isEncoded)
            {
                // Get the encoding which corresponds to the character set
                var charsetEncoding = EncodingFinder.FindEncoding(tempCharset);
                
                // Collect all adjacent encoded words with the same encoding and charset
                var contents = new List<string> { tempContent };
                
                while (i + 1 < matches.Count && 
                       matches[i + 1].IsEncoded &&
                       tempEncoding == matches[i + 1].Encoding && 
                       tempCharset == matches[i + 1].Charset)
                {
                    i++;
                    tempValue += matches[i].Value;
                    contents.Add(matches[i].Content);
                }

                // Encoding may also be written in lowercase
                switch (tempEncoding.ToUpperInvariant())
                {
                    // RFC:
                    // The "B" encoding is identical to the "BASE64" 
                    // encoding defined by RFC 2045.
                    // http://tools.ietf.org/html/rfc2045#section-6.8
                    case "B":
                        // For Base64, decode each part separately and concatenate the results
                        decodedText = string.Empty;
                        foreach (var content in contents)
                        {
                            decodedText += Base64.Decode(content, charsetEncoding);
                        }
                        break;

                    // RFC:
                    // The "Q" encoding is similar to the "Quoted-Printable" content-
                    // transfer-encoding defined in RFC 2045.
                    // There are more details to this. Please check
                    // http://tools.ietf.org/html/rfc2047#section-4.2
                    // 
                    case "Q":
                        // For Quoted-Printable, concatenate the encoded strings before decoding
                        // This is crucial for properly handling multi-byte character sequences that are split
                        var concatenatedContent = string.Join("", contents);
                        decodedText = QuotedPrintable.DecodeEncodedWord(concatenatedContent, charsetEncoding);
                        break;

                    default:
                        throw new ArgumentException($"The encoding {tempEncoding} was not recognized");
                }
            }
            else
            {
                //If the value is not encoded, use the raw value
                decodedText = tempValue;
            }

            // Replace our encoded value with our decoded value
            decodedWords += decodedText;
        }

        return decodedWords;
    }
    #endregion
}