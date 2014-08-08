using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace MsgReader.Mime.Decode
{
	/// <summary>
	/// This class is responsible for decoding parameters that has been encoded with:<br/>
	/// <list type="bullet">
	/// <item>
	///    <b>Continuation</b><br/>
	///    This is where a single parameter has such a long value that it could
	///    be wrapped while in transit. Instead multiple parameters is used on each line.<br/>
	///    <br/>
	///    <b>Example</b><br/>
	///    From: <c>Content-Type: text/html; boundary="someVeryLongStringHereWhichCouldBeWrappedInTransit"</c><br/>
	///    To: <c>Content-Type: text/html; boundary*0="someVeryLongStringHere" boundary*1="WhichCouldBeWrappedInTransit"</c><br/>
	/// </item>
	/// <item>
	///    <b>Encoding</b><br/>
	///    Sometimes other characters then ASCII characters are needed in parameters.<br/>
	///    The parameter is then given a different name to specify that it is encoded.<br/>
	///    <br/>
	///    <b>Example</b><br/>
	///    From: <c>Content-Disposition attachment; filename="specialCharsÆØÅ"</c><br/>
	///    To: <c>Content-Disposition attachment; filename*="ISO-8859-1'en-us'specialCharsC6D8C0"</c><br/>
	///    This encoding is almost the same as <see cref="EncodedWord"/> encoding, and is used to decode the value.<br/>
	/// </item>
	/// <item>
	///    <b>Continuation and Encoding</b><br/>
	///    Both Continuation and Encoding can be used on the same time.<br/>
	///    <br/>
	///    <b>Example</b><br/>
	///    From: <c>Content-Disposition attachment; filename="specialCharsÆØÅWhichIsSoLong"</c><br/>
	///    To: <c>Content-Disposition attachment; filename*0*="ISO-8859-1'en-us'specialCharsC6D8C0"; filename*1*="WhichIsSoLong"</c><br/>
	///    This could also be encoded as:<br/>
	///    To: <c>Content-Disposition attachment; filename*0*="ISO-8859-1'en-us'specialCharsC6D8C0"; filename*1="WhichIsSoLong"</c><br/>
	///    Notice that <c>filename*1</c> does not have an <c>*</c> after it - denoting it IS NOT encoded.<br/>
	///    There are some rules about this:<br/>
	///    <list type="number">
	///      <item>The encoding must be mentioned in the first part (filename*0*), which has to be encoded.</item>
	///      <item>No other part must specify an encoding, but if encoded it uses the encoding mentioned in the first part.</item>
	///      <item>Parts may be encoded or not in any order.</item>
	///    </list>
	///    <br/>
	/// </item>
	/// </list>
	/// More information and the specification is available in <see href="http://tools.ietf.org/html/rfc2231">RFC 2231</see>.
	/// </summary>
	internal static class Rfc2231Decoder
    {
        #region Decode
        /// <summary>
	    /// Decodes a string of the form:<br/>
	    /// <c>value0; key1=value1; key2=value2; key3=value3</c><br/>
	    /// The returned List of key value pairs will have the key as key and the decoded value as value.<br/>
	    /// The first value0 will have a key of <see cref="string.Empty"/>.<br/>
	    /// <br/>
	    /// If continuation is used, then multiple keys will be merged into one key with the different values
	    /// decoded into on big value for that key.<br/>
	    /// Example:<br/>
	    /// <code>
	    /// title*0=part1
	    /// title*1=part2
	    /// </code>
	    /// will have key and value of:<br></br>
	    /// <c>title=decode(part1)decode(part2)</c>
	    /// </summary>
	    /// <param name="toDecode">The string to decode.</param>
	    /// <returns>A list of decoded key value pairs.</returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="toDecode"/> is <see langword="null"/></exception>
	    public static List<KeyValuePair<string, string>> Decode(string toDecode)
	    {
	        if (toDecode == null)
	            throw new ArgumentNullException("toDecode");

	        // Normalize the input to take account for missing semicolons after parameters.
	        // Example
	        // text/plain; charset=\"iso-8859-1\" name=\"somefile.txt\" or
	        // text/plain;\tcharset=\"iso-8859-1\"\tname=\"somefile.txt\"
	        // is normalized to
	        // text/plain; charset=\"iso-8859-1\"; name=\"somefile.txt\"
	        // Only works for parameters inside quotes
	        // \s = matches whitespace
	        toDecode = Regex.Replace(toDecode, "=\\s*\"(?<value>[^\"]*)\"\\s", "=\"${value}\"; ");

	        // Normalize 
	        // Since the above only works for parameters inside quotes, we need to normalize
	        // the special case with the first parameter.
	        // Example:
	        // attachment filename="foo"
	        // is normalized to
	        // attachment; filename="foo"
	        // ^ = matches start of line (when not inside square bracets [])
	        toDecode = Regex.Replace(toDecode, @"^(?<first>[^;\s]+)\s(?<second>[^;\s]+)", "${first}; ${second}");

	        // Split by semicolon, but only if not inside quotes
	        var splitted = Utility.SplitStringWithCharNotInsideQuotes(toDecode.Trim(), ';');

	        var collection = new List<KeyValuePair<string, string>>(splitted.Count);

	        foreach (var part in splitted)
	        {
	            // Empty strings should not be processed
	            if (part.Trim().Length == 0)
	                continue;

	            var keyValue = part.Trim().Split(new[] {'='}, 2);
	            switch (keyValue.Length)
	            {
	                case 1:
	                    collection.Add(new KeyValuePair<string, string>("", keyValue[0]));
	                    break;
	                case 2:
	                    collection.Add(new KeyValuePair<string, string>(keyValue[0], keyValue[1]));
	                    break;
	                default:
	                    throw new ArgumentException("When splitting the part \"" + part + "\" by = there was " +
	                                                keyValue.Length +
	                                                " parts. Only 1 and 2 are supported");
	            }
	        }

	        return DecodePairs(collection);
	    }
	    #endregion

        #region DecodePairs
        /// <summary>
	    /// Decodes the list of key value pairs into a decoded list of key value pairs.<br/>
	    /// There may be less keys in the decoded list, but then the values for the lost keys will have been appended
	    /// to the new key.
	    /// </summary>
	    /// <param name="pairs">The pairs to decode</param>
	    /// <returns>A decoded list of pairs</returns>
	    private static List<KeyValuePair<string, string>> DecodePairs(IList<KeyValuePair<string, string>> pairs)
	    {
	        if (pairs == null)
	            throw new ArgumentNullException("pairs");

	        var resultPairs = new List<KeyValuePair<string, string>>(pairs.Count);

	        var pairsCount = pairs.Count;
	        for (var i = 0; i < pairsCount; i++)
	        {
	            var currentPair = pairs[i];
	            var key = currentPair.Key;
	            var value = Utility.RemoveQuotesIfAny(currentPair.Value);

	            // Is it a continuation parameter? (encoded or not)
	            if (key.EndsWith("*0", StringComparison.OrdinalIgnoreCase) ||
	                key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase))
	            {
	                // This encoding will not be used if we get into the if which tells us
	                // that the whole continuation is not encoded

	                var encoding = "notEncoded - Value here is never used";

	                // Now lets find out if it is encoded too.
	                if (key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase))
	                {
	                    // It is encoded.

	                    // Fetch out the encoding for later use and decode the value
	                    // If the value was not encoded as the email specified
	                    // encoding will be set to null. This will be used later.
	                    value = DecodeSingleValue(value, out encoding);

	                    // Find the right key to use to store the full value
	                    // Remove the start *0 which tells is it is a continuation, and the first one
	                    // And remove the * afterwards which tells us it is encoded
	                    key = key.Replace("*0*", "");
	                }
	                else
	                {
	                    // It is not encoded, and no parts of the continuation is encoded either

	                    // Find the right key to use to store the full value
	                    // Remove the start *0 which tells is it is a continuation, and the first one
	                    key = key.Replace("*0", "");
	                }

	                // The StringBuilder will hold the full decoded value from all continuation parts
	                var builder = new StringBuilder();

	                // Append the decoded value
	                builder.Append(value);

	                // Now go trough the next keys to see if they are part of the continuation
	                for (int j = i + 1, continuationCount = 1; j < pairsCount; j++, continuationCount++)
	                {
	                    var jKey = pairs[j].Key;
	                    var valueJKey = Utility.RemoveQuotesIfAny(pairs[j].Value);

	                    if (jKey.Equals(key + "*" + continuationCount))
	                    {
	                        // This value part of the continuation is not encoded
	                        // Therefore remove qoutes if any and add to our stringbuilder
	                        builder.Append(valueJKey);

	                        // Remember to increment i, as we have now treated one more KeyValuePair
	                        i++;
	                    }
	                    else if (jKey.Equals(key + "*" + continuationCount + "*"))
	                    {
	                        // We will not get into this part if the first part was not encoded
	                        // Therefore the encoding will only be used if and only if the
	                        // first part was encoded, in which case we have remembered the encoding used

	                        // Sometimes an email creator says that a string was encoded, but it really
	                        // `was not. This is to catch that problem.
	                        if (encoding != null)
	                        {
	                            // This value part of the continuation is encoded
	                            // the encoding is not given in the current value,
	                            // but was given in the first continuation, which we remembered for use here
	                            valueJKey = DecodeSingleValue(valueJKey, encoding);
	                        }
	                        builder.Append(valueJKey);

	                        // Remember to increment i, as we have now treated one more KeyValuePair
	                        i++;
	                    }
	                    else
	                    {
	                        // No more keys for this continuation
	                        break;
	                    }
	                }

	                // Add the key and the full value as a pair
	                value = builder.ToString();
	                resultPairs.Add(new KeyValuePair<string, string>(key, value));
	            }
	            else if (key.EndsWith("*", StringComparison.OrdinalIgnoreCase))
	            {
	                // This parameter is only encoded - it is not part of a continuation
	                // We need to change the key from "<key>*" to "<key>" and decode the value

	                // To get the key we want, we remove the last * that denotes
	                // that the value hold by the key was encoded
	                key = key.Replace("*", "");

	                // Decode the value
	                string throwAway;
	                value = DecodeSingleValue(value, out throwAway);

	                // Now input the new value with the new key
	                resultPairs.Add(new KeyValuePair<string, string>(key, value));
	            }
	            else
	            {
	                // Fully normal key - the value is not encoded
	                // Therefore nothing to do, and we can simply pass the pair
	                // as being decoded now
	                resultPairs.Add(currentPair);
	            }
	        }

	        return resultPairs;
	    }
	    #endregion

        #region DecodeSingleValue
        /// <summary>
	    /// This will decode a single value of the form: <c>ISO-8859-1'en-us'%3D%3DIamHere</c><br/>
	    /// Which is basically a <see cref="EncodedWord"/> form just using % instead of =<br/>
	    /// Notice that 'en-us' part is not used for anything.<br/>
	    /// <br/>
	    /// If the single value given is not on the correct form, it will be returned without 
	    /// being decoded and <paramref name="encodingUsed"/> will be set to <see langword="null"/>.
	    /// </summary>
	    /// <param name="encodingUsed">
	    /// The encoding used to decode with - it is given back for later use.<br/>
	    /// <see langword="null"/> if input was not in the correct form.
	    /// </param>
	    /// <param name="toDecode">The value to decode</param>
	    /// <returns>
	    /// The decoded value that corresponds to <paramref name="toDecode"/> or if
	    /// <paramref name="toDecode"/> is not on the correct form, it will be non-decoded.
	    /// </returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="toDecode"/> is <see langword="null"/></exception>
	    private static string DecodeSingleValue(string toDecode, out string encodingUsed)
	    {
	        if (toDecode == null)
	            throw new ArgumentNullException("toDecode");

	        // Check if input has a part describing the encoding
	        if (toDecode.IndexOf('\'') == -1)
	        {
	            // The input was not encoded (at least not valid) and it is returned as is
	            encodingUsed = null;
	            return toDecode;
	        }
	        encodingUsed = toDecode.Substring(0, toDecode.IndexOf('\''));
	        toDecode = toDecode.Substring(toDecode.LastIndexOf('\'') + 1);
	        return DecodeSingleValue(toDecode, encodingUsed);
	    }
	    #endregion

        #region DecodeSingleValue
        /// <summary>
	    /// This will decode a single value of the form: %3D%3DIamHere
	    /// Which is basically a <see cref="EncodedWord"/> form just using % instead of =
	    /// </summary>
	    /// <param name="valueToDecode">The value to decode</param>
	    /// <param name="encoding">The encoding used to decode with</param>
	    /// <returns>The decoded value that corresponds to <paramref name="valueToDecode"/></returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="valueToDecode"/> is <see langword="null"/></exception>
	    /// <exception cref="ArgumentNullException">If <paramref name="encoding"/> is <see langword="null"/></exception>
	    private static string DecodeSingleValue(string valueToDecode, string encoding)
	    {
	        if (valueToDecode == null)
	            throw new ArgumentNullException("valueToDecode");

	        if (encoding == null)
	            throw new ArgumentNullException("encoding");

	        // The encoding used is the same as QuotedPrintable, we only
	        // need to change % to =
	        // And otherwise make it look like the correct EncodedWord encoding
	        valueToDecode = "=?" + encoding + "?Q?" + valueToDecode.Replace("%", "=") + "?=";
	        return EncodedWord.Decode(valueToDecode);
	    }
	    #endregion
	}
}