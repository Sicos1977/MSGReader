//
// DomDocument.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2023 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Globalization;
using System.IO;
using System.Text;

// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace MsgReader.Rtf;

/// <summary>
///     RTF Document
/// </summary>
/// <remarks>
///     This type is the root of RTF Dom tree structure
/// </remarks>
internal class Document
{
    #region Fields
    /// <summary>
    ///     The default rtf encoding
    /// </summary>
    private Encoding _defaultEncoding = Encoding.Default;

    /// <summary>
    ///     Text encoding of current font
    /// </summary>
    private Encoding _fontCharSet;

    /// <summary>
    ///     text encoding of associate font
    /// </summary>
    private Encoding _associateFontCharSet;
    #endregion

    #region Constructor
    /// <summary>
    ///     initialize instance
    /// </summary>
    public Document()
    {
        Info = new DocumentInfo();
        FontTable = new Table();
        Generator = null;
    }
    #endregion

    #region Properties
    /// <summary>
    ///     Text encoding
    /// </summary>
    internal Encoding RuntimeEncoding
    {
        get
        {
            if (_fontCharSet != null)
                return _fontCharSet;

            return _associateFontCharSet ?? _defaultEncoding;
        }
    }

    /// <summary>
    ///     Font table
    /// </summary>
    private Table FontTable { get; }

    /// <summary>
    ///     Document information
    /// </summary>
    private DocumentInfo Info { get; }

    /// <summary>
    ///     Document generator
    /// </summary>
    public string Generator { get; set; }

    /// <summary>
    ///     Returns the HTML content of this RTF file
    /// </summary>
    public string HtmlContent { get; set; }
    #endregion

    #region HexValueToChar
    /// <summary>
    ///     Returns a char for the corresponding <paramref name="hexValue" />
    /// </summary>
    /// <param name="hexValue"></param>
    /// <returns></returns>
    private string HexValueToChar(string hexValue)
    {
        switch (hexValue)
        {
            case "20":
                return " ";

            case "21":
                return "!";

            case "22":
                return "\"";

            case "23":
                return "#";

            case "24":
                return "$";

            case "25":
                return "%";

            case "26":
                return "&";

            case "27":
                return "'";

            case "28":
                return "(";

            case "29":
                return "";

            case "2a":
                return "*";

            case "2b":
                return "+";

            case "2c":
                return ",";

            case "2d":
                return "-";

            case "2e":
                return ".";

            case "2f":
                return "/";

            case "30":
                return "0";

            case "31":
                return "1";

            case "32":
                return "2";

            case "33":
                return "3";

            case "34":
                return "4";

            case "35":
                return "5";

            case "36":
                return "6";

            case "37":
                return "7";

            case "38":
                return "8";

            case "39":
                return "9";

            case "3a":
                return ":";

            case "3b":
                return ";";

            case "3c":
                return "<";

            case "3d":
                return "=";

            case "3e":
                return ">";

            case "3f":
                return "?";

            case "40":
                return "@";

            case "41":
                return "A";

            case "42":
                return "B";

            case "43":
                return "C";

            case "44":
                return "D";

            case "45":
                return "E";

            case "46":
                return "F";

            case "47":
                return "G";

            case "48":
                return "H";

            case "49":
                return "I";

            case "4a":
                return "J";

            case "4b":
                return "K";

            case "4c":
                return "L";

            case "4d":
                return "M";

            case "4e":
                return "N";

            case "4f":
                return "O";

            case "50":
                return "P";

            case "51":
                return "Q";

            case "52":
                return "R";

            case "53":
                return "S";

            case "54":
                return "T";

            case "55":
                return "U";

            case "56":
                return "V";

            case "57":
                return "W";

            case "58":
                return "X";

            case "59":
                return "Y";

            case "5a":
                return "Z";

            case "5b":
                return "[";

            case "5c":
                return "\\";

            case "5d":
                return "]";

            case "5e":
                return "^";

            case "5f":
                return "_";

            case "60":
                return "`";

            case "61":
                return "a";

            case "62":
                return "b";

            case "63":
                return "c";

            case "64":
                return "d";

            case "65":
                return "e";

            case "66":
                return "f";

            case "67":
                return "g";

            case "68":
                return "h";

            case "69":
                return "i";

            case "6a":
                return "j";

            case "6b":
                return "k";

            case "6c":
                return "l";

            case "6d":
                return "m";

            case "6e":
                return "n";

            case "6f":
                return "o";

            case "70":
                return "p";

            case "71":
                return "q";

            case "72":
                return "r";

            case "73":
                return "s";

            case "74":
                return "t";

            case "75":
                return "u";

            case "76":
                return "v";

            case "77":
                return "w";

            case "78":
                return "x";

            case "79":
                return "y";

            case "7a":
                return "z";

            case "7b":
                return "{";

            case "7c":
                return "|";

            case "7d":
                return "}";

            case "7e":
                return "~";

            case "80":
                return "€";

            case "82":
                return "͵";

            case "83":
                return "ƒ";

            case "84":
                return ",,";

            case "85":
                return "...";

            case "86":
                return "†";

            case "87":
                return "‡";

            case "88":
                return "∘";

            case "89":
                return "‰";

            case "8a":
                return "Š";

            case "8b":
                return "‹";

            case "8c":
                return "Œ";

            case "8d":
                return "";

            case "8e":
                return "Ž";

            case "8f":
                return "";

            case "90":
                return "";

            case "91":
                return "‘";

            case "92":
                return "’";

            case "93":
                return "“";

            case "94":
                return "”";

            case "95":
                return "•";

            case "96":
                return "–";

            case "97":
                return "—";

            case "98":
                return "~";

            case "99":
                return "™";

            case "9a":
                return "š";

            case "9b":
                return "›";

            case "9c":
                return "œ";

            case "9e":
                return "ž";

            case "9f":
                return "Ÿ";

            case "~":
                return " ";

            case "a1":
                return "¡";

            case "a2":
                return "¢";

            case "a3":
                return "£";

            case "a4":
                return "¤";

            case "a5":
                return "¥";

            case "a6":
                return "¦";

            case "a7":
                return "§";

            case "a8":
                return "¨";

            case "a9":
                return "©";

            case "aa":
                return "ª";

            case "ab":
                return "«";

            case "ac":
                return "¬";

            case "-":
                return "-";

            case "ae":
                return "®";

            case "af":
                return "¯";

            case "b0":
                return "°";

            case "b1":
                return "±";

            case "b2":
                return "²";

            case "b3":
                return "³";

            case "b4":
                return "´";

            case "b5":
                return "µ";

            case "b6":
                return "¶";

            case "b7":
                return "·";

            case "b8":
                return "¸";

            case "b9":
                return "¹";

            case "ba":
                return "º";

            case "bb":
                return "»";

            case "bc":
                return "¼";

            case "bd":
                return "½";

            case "be":
                return "¾";

            case "bf":
                return "¿";

            case "c0":
                return "À";

            case "c1":
                return "Á";

            case "c2":
                return "Â";

            case "c3":
                return "Ã";

            case "c4":
                return "Ä";

            case "c5":
                return "Å";

            case "c6":
                return "Æ";

            case "c7":
                return "Ç";

            case "c8":
                return "È";

            case "c9":
                return "É";

            case "ca":
                return "Ê";

            case "cb":
                return "Ë";

            case "cc":
                return "Ì";

            case "cd":
                return "Í";

            case "ce":
                return "Î";

            case "cf":
                return "Ï";

            case "d0":
                return "Ð";

            case "d1":
                return "Ñ";

            case "d2":
                return "Ò";

            case "d3":
                return "Ó";

            case "d4":
                return "Ô";

            case "d5":
                return "Õ";

            case "d6":
                return "Ö";

            case "d7":
                return "×";

            case "d8":
                return "Ø";

            case "d9":
                return "Ù";

            case "da":
                return "Ú";

            case "db":
                return "Û";

            case "dc":
                return "Ü";

            case "dd":
                return "Ý";

            case "de":
                return "Þ";

            case "df":
                return "ß";

            case "e0":
                return "à";

            case "e1":
                return "á";

            case "e2":
                return "â";

            case "e3":
                return "ã";

            case "e4":
                return "ä";

            case "e5":
                return "å";

            case "e6":
                return "æ";

            case "e7":
                return "ç";

            case "e8":
                return "è";

            case "e9":
                return "é";

            case "ea":
                return "ê";

            case "eb":
                return "ë";

            case "ec":
                return "ì";

            case "ed":
                return "í";

            case "ee":
                return "î";

            case "ef":
                return "ï";

            case "f0":
                return "ð";

            case "f1":
                return "ñ";

            case "f2":
                return "ò";

            case "f3":
                return "ó";

            case "f4":
                return "ô";

            case "f5":
                return "õ";

            case "f6":
                return "ö";

            case "f7":
                return "÷";

            case "f8":
                return "ø";

            case "f9":
                return "ù";

            case "fa":
                return "ú";

            case "fb":
                return "û";

            case "fc":
                return "ü";

            case "fd":
                return "ý";

            case "fe":
                return "þ";

            case "ff":
                return "ÿ";

            default:
                return null;
        }
    }
    #endregion

    #region DeEncapsulateHtmlFromRtf
    /// <summary>
    ///     Extract HTML from the given <paramref name="rtf" />
    /// </summary>
    /// <param name="rtf">text</param>
    public void DeEncapsulateHtmlFromRtf(string rtf)
    {
        HtmlContent = null;
        var stringBuilder = new StringBuilder();
        var rtfContainsEmbeddedHtml = false;
        var hexBuffer = string.Empty;
        var ignoreText = true;

        using (var stringReader = new StringReader(rtf))
        using (var reader = new Reader(stringReader))
        {
            while (reader.ReadToken() != null)
            {
                if (reader.LastToken?.Key == "'" &&
                    reader?.Keyword != "'" &&
                    hexBuffer != string.Empty &&
                    !RuntimeEncoding.IsSingleByte)
                {
                    var value = HexValueToChar(hexBuffer);
                    if (value != null)
                    {
                        stringBuilder.Append(value);
                    }
                    else
                    {
                        // Double byte charset was detected for the last token but only one byte was used so far. 
                        // This token should carry the second byte but it doesn't.
                        // Workaround: To display it anyway, we treat it as a single byte char.
                        var buff = new[] { byte.Parse(hexBuffer, NumberStyles.HexNumber) };
                        stringBuilder.Append(RuntimeEncoding.GetString(buff));
                    }

                    hexBuffer = string.Empty;

                    if (reader.TokenType == TokenType.Text)
                    {
                        stringBuilder.Append(reader.Keyword);
                        continue;
                    }
                }

                switch (reader.TokenType)
                {
                    case TokenType.Keyword:
                        switch (reader.Keyword)
                        {
                            case Consts.Ansicpg:
                                // Read default encoding
                                _defaultEncoding = Encoding.GetEncoding(reader.Parameter);
                                break;

                            case Consts.Info:
                                // Read document information
                                ReadDocumentInfo(reader);
                                return;

                            case Consts.FromHtml:
                                rtfContainsEmbeddedHtml = true;
                                break;

                            case Consts.Fonttbl:
                                // Read font table
                                ReadFontTable(reader);
                                break;

                            case Consts.F:
                            {
                                // https://learn.microsoft.com/en-us/previous-versions/cc194829(v=msdn.10)?redirectedfrom=MSDN
                                var font = FontTable[reader.Parameter];

                                if (font != null)
                                    _fontCharSet = font.Charset == 0 ? _defaultEncoding : font.Encoding;

                                break;
                            }

                            case Consts.Af:
                            {
                                var font = FontTable[reader.Parameter];

                                if (font != null)
                                    _associateFontCharSet = font.Charset == 0 ? _defaultEncoding : font.Encoding;

                                break;
                            }

                            case Consts.HtmlRtf:

                                // \htmlrtf \htmlrtf1
                                // The de-encapsulating RTF reader MUST NOT copy any subsequent text and control words in the RTF content until
                                // the state is disabled.

                                // \htmlrtf0
                                // This control word disables an earlier instance of \htmlrtf or \htmlrtf1, thereby allowing the de-encapsulating
                                // RTF reader to evaluate subsequent text and control words in the RTF content.

                                if (reader.HasParam)
                                    ignoreText = reader.Parameter != 0;
                                else
                                    ignoreText = true;

                                break;

                            case Consts.Tab:
                                stringBuilder.Append("\t");
                                break;

                            case Consts.Lquote:
                                stringBuilder.Append("&lsquo;");
                                break;

                            case Consts.Rquote:
                                stringBuilder.Append("&rsquo;");
                                break;

                            case Consts.LdblQuote:
                                stringBuilder.Append("&ldquo;");
                                break;

                            case Consts.RdblQuote:
                                stringBuilder.Append("&rdquo;");
                                break;

                            case Consts.Bullet:
                                stringBuilder.Append("&bull;");
                                break;

                            case Consts.Endash:
                                stringBuilder.Append("&ndash;");
                                break;

                            case Consts.Emdash:
                                stringBuilder.Append("&mdash;");
                                break;

                            case Consts.Tilde:
                                stringBuilder.Append("&nbsp;");
                                break;

                            case Consts.Underscore:
                                stringBuilder.Append("&shy;");
                                break;
                        }
                        break;

                    case TokenType.ExtensionKeyword:

                        switch (reader.Keyword)
                        {
                            case Consts.HtmlTag:
                            {
                                ignoreText = false;

                                if (reader.InnerReader.Peek() == ' ')
                                    reader.InnerReader.Read();

                                var text = ReadInnerText(reader, null, true, false, true);

                                if (!string.IsNullOrEmpty(text))
                                    stringBuilder.Append(text);

                                break;
                            }
                        }

                        break;

                    case TokenType.Control:
                        switch (reader.Keyword)
                        {
                            case Consts.Apostrophe:
                                var hexValue = char.ToString((char)reader.InnerReader.Read()) + char.ToString((char)reader.InnerReader.Read());

                                // When the default encoding is a 2 byte encoding and the runTime encoding 
                                // is single byte then check if the encoded char makes any sense. If not them
                                // assume the runtime encoding is wrong and use the default encoding
                                if (!_defaultEncoding.IsSingleByte && RuntimeEncoding.IsSingleByte)
                                {
                                    var asciiValue = (int)byte.Parse(hexValue, NumberStyles.HexNumber);
                                    if (asciiValue > 127)
                                        _fontCharSet = _defaultEncoding;
                                }

                                // Convert HEX value directly when we have a single byte charset
                                if (RuntimeEncoding.IsSingleByte && _defaultEncoding.IsSingleByte)
                                {
                                    if (string.IsNullOrEmpty(hexBuffer))
                                        hexBuffer = hexValue;

                                    var buff = new[] { byte.Parse(hexBuffer, NumberStyles.HexNumber) };
                                    hexBuffer = string.Empty;
                                    stringBuilder.Append(RuntimeEncoding.GetString(buff));
                                }
                                else
                                {
                                    // If we have a double byte charset like Chinese then store the value and wait for the next HEX value
                                    if (hexBuffer == string.Empty)
                                    {
                                        hexBuffer = hexValue;
                                    }
                                    else
                                    {
                                        // Append the second HEX value and convert it 
                                        var buff = new[]
                                        {
                                            byte.Parse(hexBuffer, NumberStyles.HexNumber),
                                            byte.Parse(hexValue, NumberStyles.HexNumber)
                                        };

                                        stringBuilder.Append(RuntimeEncoding.GetString(buff));

                                        // Empty the HEX buffer 
                                        hexBuffer = string.Empty;
                                    }
                                }

                                break;

                            case Consts.U:

                                if (reader.Parameter.ToString()
                                    .StartsWith("c", StringComparison.InvariantCultureIgnoreCase))
                                    throw new Exception(
                                        "\\uc parameter not yet supported, please contact the developer on GitHub");

                                if (reader.Parameter.ToString().StartsWith("-"))
                                {
                                    // The Unicode standard permanently reserves these code point values for
                                    // UTF-16 encoding of the high and low surrogates
                                    // U+D800 to U+DFFF
                                    // 55296  -  57343

                                    var value = 65536 + int.Parse(reader.Parameter.ToString());

                                    if (value is >= 0xD800 and <= 0xDFFF)
                                    {
                                        if (!reader.ParsingHighLowSurrogate)
                                        {
                                            reader.ParsingHighLowSurrogate = true;
                                            reader.HighSurrogateValue = value;
                                        }
                                        else
                                        {
                                            var combined = ((reader.HighSurrogateValue - 0xD800) << 10) + (value - 0xDC00) + 0x10000;
                                            stringBuilder.Append($"&#{combined};");
                                            reader.ParsingHighLowSurrogate = false;
                                            reader.HighSurrogateValue = null;
                                        }
                                    }
                                    else
                                    {
                                        reader.ParsingHighLowSurrogate = false;
                                        stringBuilder.Append($"&#{value};");
                                    }
                                }
                                else
                                    stringBuilder.Append($"&#{reader.Parameter};");

                                break;
                        }

                        break;

                    case TokenType.Text:
                        if (!ignoreText)
                            stringBuilder.Append(reader.Keyword);
                        break;

                    case TokenType.None:
                    case TokenType.GroupStart:
                    case TokenType.GroupEnd:
                    case TokenType.Eof:
                        break;
                    
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        if (rtfContainsEmbeddedHtml)
            HtmlContent = stringBuilder.ToString();
    }
    #endregion

    #region ReadFontTable
    /// <summary>
    ///     Reads the font table
    /// </summary>
    /// <param name="reader"></param>
    private void ReadFontTable(Reader reader)
    {
        FontTable.Clear();

        while (reader.ReadToken() != null)
        {
            if (reader.TokenType == TokenType.GroupEnd)
                break;

            if (reader.TokenType != TokenType.GroupStart) continue;

            var index = -1;
            string name = null;
            var charset = 1;
            var nilFlag = false;

            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == TokenType.GroupEnd)
                    break;

                if (reader.TokenType == TokenType.GroupStart)
                {
                    // If we meet nested levels, then ignore
                    reader.ReadToken();
                    reader.ReadToEndOfGroup();
                    reader.ReadToken();
                }
                else
                {
                    switch (reader.Keyword)
                    {
                        case "f" when reader.HasParam:
                            index = reader.Parameter;
                            break;

                        case "fnil":
#if (WINDOWS)
                            name = SystemFonts.DefaultFont.Name;
#else
                            name = "Arial";
#endif
                            nilFlag = true;
                            break;

                        case Consts.Fcharset:
                            charset = reader.Parameter;
                            break;

                        default:
                            if (reader.CurrentToken.IsTextToken)
                            {
                                name = ReadInnerText(reader, reader.CurrentToken, false, false, false);

                                if (name != null)
                                {
                                    name = name.Trim();

                                    if (name.EndsWith(";"))
                                        name = name.Substring(0, name.Length - 1);
                                }
                            }

                            break;
                    }
                }
            }

            if (index < 0 || name == null) continue;

            if (name.EndsWith(";"))
                name = name.Substring(0, name.Length - 1);

            name = name.Trim();

            if (string.IsNullOrEmpty(name))
            {
#if (WINDOWS)
                    name = SystemFonts.DefaultFont.Name;
#else
                name = "Arial";
#endif
            }

            var font = new Font(index, name) { Charset = charset, NilFlag = nilFlag };
            FontTable.Add(font);
        }
    }
    #endregion

    #region ReadDocumentInfo
    /// <summary>
    ///     Read document information
    /// </summary>
    /// <param name="reader"></param>
    private void ReadDocumentInfo(Reader reader)
    {
        Info.Clear();
        var level = 0;

        while (reader.ReadToken() != null)
            if (reader.TokenType == TokenType.GroupStart)
            {
                level++;
            }
            else if (reader.TokenType == TokenType.GroupEnd)
            {
                level--;
                if (level < 0)
                    break;
            }
            else
            {
                switch (reader.Keyword)
                {
                    case "creatim":
                        Info.CreationTime = ReadDateTime(reader);
                        level--;
                        break;

                    case "revtim":
                        Info.RevisionTime = ReadDateTime(reader);
                        level--;
                        break;

                    case "printim":
                        Info.PrintTime = ReadDateTime(reader);
                        level--;
                        break;

                    case "buptim":
                        Info.BackupTime = ReadDateTime(reader);
                        level--;
                        break;

                    default:
                        if (reader.Keyword != null)
                            Info.SetInfo(reader.Keyword,
                                reader.HasParam
                                    ? reader.Parameter.ToString(CultureInfo.InvariantCulture)
                                    : ReadInnerText(reader, true));
                        break;
                }
            }
    }
    #endregion

    #region ReadDateTime
    /// <summary>
    ///     Read datetime
    /// </summary>
    /// <param name="reader">reader</param>
    /// <returns>datetime value</returns>
    private DateTime ReadDateTime(Reader reader)
    {
        var year = 1900;
        var month = 1;
        var day = 1;
        var hour = 0;
        var min = 0;
        var sec = 0;

        while (reader.ReadToken() != null)
        {
            if (reader.TokenType == TokenType.GroupEnd)
                break;

            switch (reader.Keyword)
            {
                case "yr":
                    year = reader.Parameter;
                    break;

                case "mo":
                    month = reader.Parameter;
                    break;

                case "dy":
                    day = reader.Parameter;
                    break;

                case "hr":
                    hour = reader.Parameter;
                    break;

                case "min":
                    min = reader.Parameter;
                    break;

                case "sec":
                    sec = reader.Parameter;
                    break;
            }
        }

        return new DateTime(year, month, day, hour, min, sec);
    }
    #endregion

    #region ReadInnerText
    /// <summary>
    ///     Read the following plain text in the current level
    /// </summary>
    /// <param name="reader">RTF reader</param>
    /// <param name="deeply">whether read the text in the sub level</param>
    private string ReadInnerText(Reader reader, bool deeply)
    {
        return ReadInnerText(reader, null, deeply, false, false);
    }

    /// <summary>
    ///     Read the following plain text in the current level
    /// </summary>
    /// <param name="reader">RTF reader</param>
    /// <param name="firstToken"></param>
    /// <param name="deeply">whether read the text in the sub level</param>
    /// <param name="breakMeetControlWord"></param>
    /// <param name="htmlExtraction">When <c>true</c> then we are de-encapsulating HTML from RTF</param>
    /// <returns>text</returns>
    private string ReadInnerText(
        Reader reader,
        Token firstToken,
        bool deeply,
        bool breakMeetControlWord,
        bool htmlExtraction)
    {
        var level = 0;
        var container = new TextContainer(this);
        container.Accept(firstToken, reader);

        while (true)
        {
            var type = reader.PeekTokenType();

            if (type == TokenType.Eof)
                break;

            if (type == TokenType.GroupStart)
            {
                level++;
            }
            else if (type == TokenType.GroupEnd)
            {
                level--;
                if (level < 0)
                    break;
            }

            reader.ReadToken();

            if (!deeply && level != 0)
                continue;

            if (htmlExtraction && reader.Keyword == Consts.Par)
            {
                container.Append(Environment.NewLine);
                continue;
            }

            container.Accept(reader.CurrentToken, reader);

            if (breakMeetControlWord)
                break;
        }

        return container.Text;
    }
    #endregion
}