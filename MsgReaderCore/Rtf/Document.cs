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
    ///     Either found through the font or language tag
    /// </summary>
    private Encoding _runtimeEncoding;
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
    internal Encoding RuntimeEncoding => _runtimeEncoding ?? _defaultEncoding;

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
        ByteBuffer byteBuffer = new();
        var ignore = true;

        using (var stringReader = new StringReader(rtf))
        using (var reader = new Reader(stringReader))
        {
            while (reader.ReadToken() != null)
            {
                if (byteBuffer.Count > 0 && reader.TokenType != TokenType.EncodedChar)
                {
                    stringBuilder.Append(byteBuffer.GetString(RuntimeEncoding));
                    byteBuffer.Clear();
                }

                switch (reader.TokenType)
                {
                    case TokenType.Keyword:
                        switch (reader.Keyword)
                        {
                            case Consts.Ansicpg:
                                // Read default encoding
                                _defaultEncoding = Font.EncodingFromCodePage(reader.Parameter);
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
                            case Consts.Af:
                            {
                                var font = FontTable[reader.Parameter];
                                _runtimeEncoding =  font.Encoding;

                                break;
                            }

                            case Consts.Lang:
                            {
                                try
                                {
                                    var lang = reader.Parameter;
                                    var culture = CultureInfo.GetCultureInfo(lang);
                                    _runtimeEncoding = Encoding.GetEncoding(culture.TextInfo.ANSICodePage);
                                }
                                catch
                                {
                                    _runtimeEncoding = _defaultEncoding;
                                }
                                break;
                            }
                                
                            case Consts.Pntxtb:
                            case Consts.Pntext:
                                if (ignore) continue;
                                reader.ReadToEndOfGroup();
                                break;

                            case Consts.HtmlRtf:

                                // \htmlrtf \htmlrtf1
                                // The de-encapsulating RTF reader MUST NOT copy any subsequent text and control words in the RTF content until
                                // the state is disabled.

                                // \htmlrtf0
                                // This control word disables an earlier instance of \htmlrtf or \htmlrtf1, thereby allowing the de-encapsulating
                                // RTF reader to evaluate subsequent text and control words in the RTF content.

                                if (reader.HasParam && reader.Parameter == 0)
                                    ignore = false;
                                else
                                    ignore = true;

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

                    case TokenType.Extension:

                        switch (reader.Keyword)
                        {
                            case Consts.HtmlTag:
                            {
                                ignore = false;

                                if (reader.InnerReader.Peek() == ' ')
                                    reader.InnerReader.Read();

                                var text = ReadInnerText(reader, null, true, true, RuntimeEncoding);

                                if (!string.IsNullOrEmpty(text))
                                    stringBuilder.Append(text);

                                break;
                            }
                        }

                        break;

                    case TokenType.EncodedChar:
                        
                        if (ignore) continue;

                        switch (reader.Keyword)
                        {
                            case Consts.Apostrophe:

                                byteBuffer.Add((byte)reader.Parameter);
                                break;

                            case Consts.U:

                                if (reader.Parameter.ToString().StartsWith("c", StringComparison.InvariantCultureIgnoreCase))
                                    throw new Exception("\\uc parameter not yet supported, please contact the developer on GitHub");

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

                        if (!ignore)
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

            var font = new Font { Charset = 1 };

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
                        case Consts.F when reader.HasParam:
                            font.Index = reader.Parameter;
                            break;

                        case Consts.Fnil:
#if (WINDOWS)
                            font.Name = SystemFonts.DefaultFont.Name;
#else
                            font.Name = "Arial";
#endif
                            break;

                        case Consts.Fcharset:
                            font.Charset = reader.Parameter;
                            break;
                        
                        default:

                            if (reader.CurrentToken.IsTextToken)
                                font.Name = ReadInnerText(reader, reader.CurrentToken, false, false, font.Encoding ?? RuntimeEncoding);

                            break;
                    }
                }
            }


            font.Name = font.Name.TrimEnd(';').Trim('\"');

            if (string.IsNullOrEmpty(font.Name))
            {
#if (WINDOWS)
                    name = SystemFonts.DefaultFont.Name;
#else
                font.Name = "Arial";
#endif
            }

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
        return ReadInnerText(reader, null, deeply, false, RuntimeEncoding);
    }

    /// <summary>
    ///     Read the following plain text in the current level
    /// </summary>
    /// <param name="reader">RTF reader</param>
    /// <param name="firstToken"></param>
    /// <param name="deeply">whether read the text in the sub level</param>
    /// <param name="htmlExtraction">When true then we are extracting HTML</param>
    /// <param name="encoding"></param>
    /// <returns>text</returns>
    private string ReadInnerText(
        Reader reader,
        Token firstToken,
        bool deeply,
        bool htmlExtraction,
        Encoding encoding)
    {
        var level = 0;
        var container = new TextContainer(encoding);
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
        }

        return container.Text;
    }
    #endregion
}