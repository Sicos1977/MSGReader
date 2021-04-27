//
// DomDocument.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2021 Magic-Sessions. (www.magic-sessions.com)
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
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web;
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace MsgReader.Rtf
{
	/// <summary>
	/// RTF Document
	/// </summary>
	/// <remarks>
	/// This type is the root of RTF Dom tree structure
	/// </remarks>
	internal class DomDocument
	{
        #region Fields
        /// <summary>
        /// The default rtf encoding
        /// </summary>
        private Encoding _defaultEncoding = Encoding.Default;

        /// <summary>
        /// Text encoding of current font
        /// </summary>
        private Encoding _fontChartset;

        /// <summary>
        /// text encoding of associate font 
        /// </summary>
        private Encoding _associateFontChartset;
        #endregion

        #region Constructor
        /// <summary>
        /// initialize instance
        /// </summary>
        public DomDocument()
        {
	        Info = new DocumentInfo();
	        FontTable = new Table();
	        Generator = null;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Text encoding
        /// </summary>
        internal Encoding RuntimeEncoding
        {
            get
            {
                if (_fontChartset != null)
                    return _fontChartset;

                return _associateFontChartset ?? _defaultEncoding;
            }
        }

        /// <summary>
        /// Font table
        /// </summary>
        private Table FontTable { get; }

        /// <summary>
        /// Document information
        /// </summary>
        private DocumentInfo Info { get; }

        /// <summary>
        /// Document generator
        /// </summary>
        public string Generator { get; set; }

        /// <summary>
        /// Format converter
        /// </summary>
        public string FormatConverter { get; set; }
        
        /// <summary>
        /// Returns the HTNL content of this RTF file
        /// </summary>
        public string HtmlContent { get; set; }
        #endregion

        #region LoadRtfText
        /// <summary>
        /// Load a rtf document from a string in rtf format and parse content
        /// </summary>
        /// <param name="rtfText">text</param>
        public void LoadRtfText(string rtfText)
        {
            HtmlContent = null;

            using(var stringReader = new StringReader(rtfText))
            using (var reader = new Reader(stringReader))
            {
                while (reader.ReadToken() != null)
                {
                    if (reader.TokenType == RtfTokenType.Control
                        || reader.TokenType == RtfTokenType.Keyword
                        || reader.TokenType == RtfTokenType.ExtKeyword)
                    {
                        switch (reader.Keyword)
                        {
                            case Consts.FromHtml:
                                // Extract html from rtf
                                ReadHtmlContent(reader);
                                return;

                            case Consts.Ansicpg:
                                // Read default encoding
                                _defaultEncoding = Encoding.GetEncoding(reader.Parameter);
                                break;

                            case Consts.Fonttbl:
                                // Read font table
                                ReadFontTable(reader);
                                break;

                            case Consts.F:
                                try
                                {
                                    _fontChartset = FontTable[reader.Parameter].Encoding;
                                }
                                catch
                                {
                                    _fontChartset = Encoding.Default;
                                }

                                break;

                            case Consts.Af:
                                _associateFontChartset = FontTable[reader.Parameter].Encoding;
                                break;

                            case Consts.Generator:
                                // Read document generator
                                Generator = ReadInnerText(reader, true);
                                break;

                            case Consts.Info:
                                // Read document information
                                ReadDocumentInfo(reader);
                                return;

                            case Consts.Xmlns:
                                // Unsupport xml namespace
                                ReadToEndOfGroup(reader);
                                break;

                            default:
                                // Unsupport keyword
                                if (reader.TokenType == RtfTokenType.ExtKeyword && reader.FirstTokenInGroup)
                                {
                                    // If we have an unsupport extern keyword , and this token is the first token in 
                                    // then current group , then ingore the whole group.
                                    ReadToEndOfGroup(reader);
                                }

                                break;
                        }
                    }
                }
            }
        }
        #endregion

        #region ReadToEndOfGroup
        /// <summary>
        /// Read and ignore data , until just the end of the current group, preserve the end.
        /// </summary>
        /// <param name="reader"></param>
        private void ReadToEndOfGroup(Reader reader)
        {
	        reader.ReadToEndOfGroup();
        }
        #endregion

        #region ReadFontTable
        /// <summary>
        /// Read font table
        /// </summary>
        /// <param name="reader"></param>
        private void ReadFontTable(Reader reader)
        {
	        FontTable.Clear();
	        while (reader.ReadToken() != null)
	        {
		        if (reader.TokenType == RtfTokenType.GroupEnd)
			        break;

		        if (reader.TokenType == RtfTokenType.GroupStart)
		        {
			        var index = -1;
			        string name = null;
			        var charset = 1;
			        var nilFlag = false;
			        while (reader.ReadToken() != null)
			        {
				        if (reader.TokenType == RtfTokenType.GroupEnd)
					        break;

				        if (reader.TokenType == RtfTokenType.GroupStart)
				        {
					        // if meet nested level , then ignore
					        reader.ReadToken();
					        ReadToEndOfGroup(reader);
					        reader.ReadToken();
				        }
				        else if (reader.Keyword == "f" && reader.HasParam)
					        index = reader.Parameter;
				        else if (reader.Keyword == "fnil")
				        {
					        name = SystemFonts.DefaultFont.Name;
					        nilFlag = true;
				        }
				        else if (reader.Keyword == Consts.Fcharset)
					        charset = reader.Parameter;
				        else if (reader.CurrentToken.IsTextToken)
				        {
					        name = ReadInnerText(
						        reader,
						        reader.CurrentToken,
						        false,
						        false,
						        false);

					        if (name != null)
					        {
						        name = name.Trim();

						        if (name.EndsWith(";"))
							        name = name.Substring(0, name.Length - 1);
					        }
				        }
			        }

			        if (index >= 0 && name != null)
			        {
				        if (name.EndsWith(";"))
					        name = name.Substring(0, name.Length - 1);

				        name = name.Trim();
				        if (string.IsNullOrEmpty(name))
					        name = SystemFonts.DefaultFont.Name;

				        var font = new Font(index, name) { Charset = charset, NilFlag = nilFlag };
				        FontTable.Add(font);
			        }
		        }
	        }
        }
        #endregion

        #region ReadDocumentInfo
        /// <summary>
        /// Read document information
        /// </summary>
        /// <param name="reader"></param>
        private void ReadDocumentInfo(Reader reader)
        {
	        Info.Clear();
	        var level = 0;
	        while (reader.ReadToken() != null)
	        {
		        if (reader.TokenType == RtfTokenType.GroupStart)
			        level++;
		        else if (reader.TokenType == RtfTokenType.GroupEnd)
		        {
			        level--;
			        if (level < 0)
			        {
				        break;
			        }
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
							        reader.HasParam ? reader.Parameter.ToString(CultureInfo.InvariantCulture) : ReadInnerText(reader, true));
					        break;
			        }
		        }
	        }
        }
        #endregion

        #region ReadDateTime
        /// <summary>
        /// Read datetime
        /// </summary>
        /// <param name="reader">reader</param>
        /// <returns>datetime value</returns>
        private DateTime ReadDateTime(Reader reader)
        {
	        var yr = 1900;
	        var mo = 1;
	        var dy = 1;
	        var hr = 0;
	        var min = 0;
	        var sec = 0;
	        while (reader.ReadToken() != null)
	        {
		        if (reader.TokenType == RtfTokenType.GroupEnd)
		        {
			        break;
		        }
		        switch (reader.Keyword)
		        {
			        case "yr":
				        yr = reader.Parameter;
				        break;

			        case "mo":
				        mo = reader.Parameter;
				        break;

			        case "dy":
				        dy = reader.Parameter;
				        break;

			        case "hr":
				        hr = reader.Parameter;
				        break;

			        case "min":
				        min = reader.Parameter;
				        break;

			        case "sec":
				        sec = reader.Parameter;
				        break;
		        }
	        }
	        return new DateTime(yr, mo, dy, hr, min, sec);
        }
        #endregion

        #region ReadInnerText
        /// <summary>
        /// Read the following plain text in the current level
        /// </summary>
        /// <param name="reader">RTF reader</param>
        /// <param name="deeply">whether read the text in the sub level</param>
        private string ReadInnerText(Reader reader, bool deeply)
        {
	        return ReadInnerText(
		        reader,
		        null,
		        deeply,
		        false,
		        false);
        }

        /// <summary>
        /// Read the following plain text in the current level
        /// </summary>
        /// <param name="reader">RTF reader</param>
        /// <param name="firstToken"></param>
        /// <param name="deeply">whether read the text in the sub level</param>
        /// <param name="breakMeetControlWord"></param>
        /// <param name="htmlMode"></param>
        /// <returns>text</returns>
        private string ReadInnerText(
	        Reader reader,
	        Token firstToken,
	        bool deeply,
	        bool breakMeetControlWord,
	        bool htmlMode)
        {
	        var level = 0;
	        var container = new TextContainer(this);
	        container.Accept(firstToken, reader);

	        while (true)
	        {
		        var type = reader.PeekTokenType();

		        if (type == RtfTokenType.Eof)
			        break;

                if (type == RtfTokenType.GroupStart)
                    level++;
                else if (type == RtfTokenType.GroupEnd)
                {
                    level--;
                    if (level < 0)
                        break;
                }

                reader.ReadToken();

		        if (!deeply && level != 0) continue;

		        if (htmlMode && reader.Keyword == Consts.Par)
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

        #region ReadHtmlContent
        /// <summary>
        /// Read embedded Html content from rtf
        /// </summary>
        /// <param name="reader"></param>
        private void ReadHtmlContent(Reader reader)
        {
	        var stringBuilder = new StringBuilder();
	        var htmlState = true;
	        var hexBuffer = string.Empty;
            int? fontIndex = null;
            var encoding = _defaultEncoding;

            while (reader.ReadToken() != null)
	        {
                if (reader.LastToken?.Key == "'" && reader?.Keyword != "'" && hexBuffer != string.Empty && !encoding.IsSingleByte)
                {
                    #region Collapse
                    switch(hexBuffer)
                    {
                        case "20": stringBuilder.Append(" "); break;
                        case "21": stringBuilder.Append("!"); break;
                        case "22": stringBuilder.Append("\""); break;
                        case "23": stringBuilder.Append("#"); break;
                        case "24": stringBuilder.Append("$"); break;
                        case "25": stringBuilder.Append("%"); break;
                        case "26": stringBuilder.Append("&"); break;
                        case "27": stringBuilder.Append("'"); break;
                        case "28": stringBuilder.Append("("); break;
                        case "29": stringBuilder.Append(")"); break;
                        case "2a": stringBuilder.Append("*"); break;
                        case "2b": stringBuilder.Append("+"); break;
                        case "2c": stringBuilder.Append(","); break;
                        case "2d": stringBuilder.Append("-"); break;
                        case "2e": stringBuilder.Append("."); break;
                        case "2f": stringBuilder.Append("/"); break;
                        case "30": stringBuilder.Append("0"); break;
                        case "31": stringBuilder.Append("1"); break;
                        case "32": stringBuilder.Append("2"); break;
                        case "33": stringBuilder.Append("3"); break;
                        case "34": stringBuilder.Append("4"); break;
                        case "35": stringBuilder.Append("5"); break;
                        case "36": stringBuilder.Append("6"); break;
                        case "37": stringBuilder.Append("7"); break;
                        case "38": stringBuilder.Append("8"); break;
                        case "39": stringBuilder.Append("9"); break;
                        case "3a": stringBuilder.Append(":"); break;
                        case "3b": stringBuilder.Append(";"); break;
                        case "3c": stringBuilder.Append("<"); break;
                        case "3d": stringBuilder.Append("="); break;
                        case "3e": stringBuilder.Append(">"); break;
                        case "3f": stringBuilder.Append("?"); break;
                        case "40": stringBuilder.Append("@"); break;
                        case "41": stringBuilder.Append("A"); break;
                        case "42": stringBuilder.Append("B"); break;
                        case "43": stringBuilder.Append("C"); break;
                        case "44": stringBuilder.Append("D"); break;
                        case "45": stringBuilder.Append("E"); break;
                        case "46": stringBuilder.Append("F"); break;
                        case "47": stringBuilder.Append("G"); break;
                        case "48": stringBuilder.Append("H"); break;
                        case "49": stringBuilder.Append("I"); break;
                        case "4a": stringBuilder.Append("J"); break;
                        case "4b": stringBuilder.Append("K"); break;
                        case "4c": stringBuilder.Append("L"); break;
                        case "4d": stringBuilder.Append("M"); break;
                        case "4e": stringBuilder.Append("N"); break;
                        case "4f": stringBuilder.Append("O"); break;
                        case "50": stringBuilder.Append("P"); break;
                        case "51": stringBuilder.Append("Q"); break;
                        case "52": stringBuilder.Append("R"); break;
                        case "53": stringBuilder.Append("S"); break;
                        case "54": stringBuilder.Append("T"); break;
                        case "55": stringBuilder.Append("U"); break;
                        case "56": stringBuilder.Append("V"); break;
                        case "57": stringBuilder.Append("W"); break;
                        case "58": stringBuilder.Append("X"); break;
                        case "59": stringBuilder.Append("Y"); break;
                        case "5a": stringBuilder.Append("Z"); break;
                        case "5b": stringBuilder.Append("["); break;
                        case "5c": stringBuilder.Append("\\"); break;
                        case "5d": stringBuilder.Append("]"); break;
                        case "5e": stringBuilder.Append("^"); break;
                        case "5f": stringBuilder.Append("_"); break;
                        case "60": stringBuilder.Append("`"); break;
                        case "61": stringBuilder.Append("a"); break;
                        case "62": stringBuilder.Append("b"); break;
                        case "63": stringBuilder.Append("c"); break;
                        case "64": stringBuilder.Append("d"); break;
                        case "65": stringBuilder.Append("e"); break;
                        case "66": stringBuilder.Append("f"); break;
                        case "67": stringBuilder.Append("g"); break;
                        case "68": stringBuilder.Append("h"); break;
                        case "69": stringBuilder.Append("i"); break;
                        case "6a": stringBuilder.Append("j"); break;
                        case "6b": stringBuilder.Append("k"); break;
                        case "6c": stringBuilder.Append("l"); break;
                        case "6d": stringBuilder.Append("m"); break;
                        case "6e": stringBuilder.Append("n"); break;
                        case "6f": stringBuilder.Append("o"); break;
                        case "70": stringBuilder.Append("p"); break;
                        case "71": stringBuilder.Append("q"); break;
                        case "72": stringBuilder.Append("r"); break;
                        case "73": stringBuilder.Append("s"); break;
                        case "74": stringBuilder.Append("t"); break;
                        case "75": stringBuilder.Append("u"); break;
                        case "76": stringBuilder.Append("v"); break;
                        case "77": stringBuilder.Append("w"); break;
                        case "78": stringBuilder.Append("x"); break;
                        case "79": stringBuilder.Append("y"); break;
                        case "7a": stringBuilder.Append("z"); break;
                        case "7b": stringBuilder.Append("{"); break;
                        case "7c": stringBuilder.Append("|"); break;
                        case "7d": stringBuilder.Append("}"); break;
                        case "7e": stringBuilder.Append("~"); break;
                        case "80": stringBuilder.Append("€"); break;
                        case "82": stringBuilder.Append("͵"); break;
                        case "83": stringBuilder.Append("ƒ"); break;
                        case "84": stringBuilder.Append(",,"); break;
                        case "85": stringBuilder.Append("..."); break;
                        case "86": stringBuilder.Append("†"); break;
                        case "87": stringBuilder.Append("‡"); break;
                        case "88": stringBuilder.Append("∘"); break;
                        case "89": stringBuilder.Append("‰"); break;
                        case "8a": stringBuilder.Append("Š"); break;
                        case "8b": stringBuilder.Append("‹"); break;
                        case "8c": stringBuilder.Append("Œ"); break;
                        case "8d": stringBuilder.Append(""); break;
                        case "8e": stringBuilder.Append("Ž"); break;
                        case "8f": stringBuilder.Append(""); break;
                        case "90": stringBuilder.Append(""); break;
                        case "91": stringBuilder.Append("‘"); break;
                        case "92": stringBuilder.Append("’"); break;
                        case "93": stringBuilder.Append("“"); break;
                        case "94": stringBuilder.Append("”"); break;
                        case "95": stringBuilder.Append("•"); break;
                        case "96": stringBuilder.Append("–"); break;
                        case "97": stringBuilder.Append("—"); break;
                        case "98": stringBuilder.Append("~"); break;
                        case "99": stringBuilder.Append("™"); break;
                        case "9a": stringBuilder.Append("š"); break;
                        case "9b": stringBuilder.Append("›"); break;
                        case "9c": stringBuilder.Append("œ"); break;
                        case "9e": stringBuilder.Append("ž"); break;
                        case "9f": stringBuilder.Append("Ÿ"); break;
                        case "~":  stringBuilder.Append(" "); break;
                        case "a1": stringBuilder.Append("¡"); break;
                        case "a2": stringBuilder.Append("¢"); break;
                        case "a3": stringBuilder.Append("£"); break;
                        case "a4": stringBuilder.Append("¤"); break;
                        case "a5": stringBuilder.Append("¥"); break;
                        case "a6": stringBuilder.Append("¦"); break;
                        case "a7": stringBuilder.Append("§"); break;
                        case "a8": stringBuilder.Append("¨"); break;
                        case "a9": stringBuilder.Append("©"); break;
                        case "aa": stringBuilder.Append("ª"); break;
                        case "ab": stringBuilder.Append("«"); break;
                        case "ac": stringBuilder.Append("¬"); break;
                        case "-" : stringBuilder.Append("-"); break;
                        case "ae": stringBuilder.Append("®"); break;
                        case "af": stringBuilder.Append("¯"); break;
                        case "b0": stringBuilder.Append("°"); break;
                        case "b1": stringBuilder.Append("±"); break;
                        case "b2": stringBuilder.Append("²"); break;
                        case "b3": stringBuilder.Append("³"); break;
                        case "b4": stringBuilder.Append("´"); break;
                        case "b5": stringBuilder.Append("µ"); break;
                        case "b6": stringBuilder.Append("¶"); break;
                        case "b7": stringBuilder.Append("·"); break;
                        case "b8": stringBuilder.Append("¸"); break;
                        case "b9": stringBuilder.Append("¹"); break;
                        case "ba": stringBuilder.Append("º"); break;
                        case "bb": stringBuilder.Append("»"); break;
                        case "bc": stringBuilder.Append("¼"); break;
                        case "bd": stringBuilder.Append("½"); break;
                        case "be": stringBuilder.Append("¾"); break;
                        case "bf": stringBuilder.Append("¿"); break;
                        case "c0": stringBuilder.Append("À"); break;
                        case "c1": stringBuilder.Append("Á"); break;
                        case "c2": stringBuilder.Append("Â"); break;
                        case "c3": stringBuilder.Append("Ã"); break;
                        case "c4": stringBuilder.Append("Ä"); break;
                        case "c5": stringBuilder.Append("Å"); break;
                        case "c6": stringBuilder.Append("Æ"); break;
                        case "c7": stringBuilder.Append("Ç"); break;
                        case "c8": stringBuilder.Append("È"); break;
                        case "c9": stringBuilder.Append("É"); break;
                        case "ca": stringBuilder.Append("Ê"); break;
                        case "cb": stringBuilder.Append("Ë"); break;
                        case "cc": stringBuilder.Append("Ì"); break;
                        case "cd": stringBuilder.Append("Í"); break;
                        case "ce": stringBuilder.Append("Î"); break;
                        case "cf": stringBuilder.Append("Ï"); break;
                        case "d0": stringBuilder.Append("Ð"); break;
                        case "d1": stringBuilder.Append("Ñ"); break;
                        case "d2": stringBuilder.Append("Ò"); break;
                        case "d3": stringBuilder.Append("Ó"); break;
                        case "d4": stringBuilder.Append("Ô"); break;
                        case "d5": stringBuilder.Append("Õ"); break;
                        case "d6": stringBuilder.Append("Ö"); break;
                        case "d7": stringBuilder.Append("×"); break;
                        case "d8": stringBuilder.Append("Ø"); break;
                        case "d9": stringBuilder.Append("Ù"); break;
                        case "da": stringBuilder.Append("Ú"); break;
                        case "db": stringBuilder.Append("Û"); break;
                        case "dc": stringBuilder.Append("Ü"); break;
                        case "dd": stringBuilder.Append("Ý"); break;
                        case "de": stringBuilder.Append("Þ"); break;
                        case "df": stringBuilder.Append("ß"); break;
                        case "e0": stringBuilder.Append("à"); break;
                        case "e1": stringBuilder.Append("á"); break;
                        case "e2": stringBuilder.Append("â"); break;
                        case "e3": stringBuilder.Append("ã"); break;
                        case "e4": stringBuilder.Append("ä"); break;
                        case "e5": stringBuilder.Append("å"); break;
                        case "e6": stringBuilder.Append("æ"); break;
                        case "e7": stringBuilder.Append("ç"); break;
                        case "e8": stringBuilder.Append("è"); break;
                        case "e9": stringBuilder.Append("é"); break;
                        case "ea": stringBuilder.Append("ê"); break;
                        case "eb": stringBuilder.Append("ë"); break;
                        case "ec": stringBuilder.Append("ì"); break;
                        case "ed": stringBuilder.Append("í"); break;
                        case "ee": stringBuilder.Append("î"); break;
                        case "ef": stringBuilder.Append("ï"); break;
                        case "f0": stringBuilder.Append("ð"); break;
                        case "f1": stringBuilder.Append("ñ"); break;
                        case "f2": stringBuilder.Append("ò"); break;
                        case "f3": stringBuilder.Append("ó"); break;
                        case "f4": stringBuilder.Append("ô"); break;
                        case "f5": stringBuilder.Append("õ"); break;
                        case "f6": stringBuilder.Append("ö"); break;
                        case "f7": stringBuilder.Append("÷"); break;
                        case "f8": stringBuilder.Append("ø"); break;
                        case "f9": stringBuilder.Append("ù"); break;
                        case "fa": stringBuilder.Append("ú"); break;
                        case "fb": stringBuilder.Append("û"); break;
                        case "fc": stringBuilder.Append("ü"); break;
                        case "fd": stringBuilder.Append("ý"); break;
                        case "fe": stringBuilder.Append("þ"); break;
                        case "ff": stringBuilder.Append("ÿ"); break;

                        default:
                            // Double byte charset was detected for the last token but only one byte was used so far. 
                            // This token should carry the second byte but it doesn't.
                            // Workaround: To display it anyway, we treat it as a single byte char.
                            var buff = new[] { byte.Parse(hexBuffer, NumberStyles.HexNumber) };
                            stringBuilder.Append(encoding.GetString(buff)); break;
                    }
                    #endregion

                    hexBuffer = string.Empty;

                    if (reader.TokenType == RtfTokenType.Text)
                    {
                        stringBuilder.Append(reader.Keyword);
                        continue;
                    }
                }

                switch (reader.Keyword)
		        {
                    case Consts.Generator:
                        // Read document generator
                        Generator = ReadInnerText(reader, true);
                        break;

                    case Consts.FormatConverter:
                        FormatConverter = ReadInnerText(reader, true);
                        break;

                    case Consts.Fonttbl:
                        // Read font table
                        ReadFontTable(reader);
                        break;

                    case Consts.F:
                        if (reader.TokenType == RtfTokenType.Text)
                            goto default;

                        if (reader.HasParam)
                            fontIndex = reader.Parameter;
                        break;

                    case Consts.HtmlRtf:
                        if (!reader.HasParam)
                            htmlState = true;
                        if (reader.HasParam && reader.Parameter == 0)
                            htmlState = false;
                        break;

			        case Consts.MHtmlTag:
				        if (reader.HasParam && reader.Parameter == 0)
				            htmlState = false;
				        else
				        {
				            if (hexBuffer != string.Empty)
				            {
				                var buff = new[] {byte.Parse(hexBuffer, NumberStyles.HexNumber)};
				                hexBuffer = string.Empty;
				                stringBuilder.Append(encoding.GetString(buff));
                                htmlState = true;
				            }
                            else
                                htmlState = false;
				        }
				        break;

                    case Consts.HtmlTag:
                    {
                        if (reader.InnerReader.Peek() == ' ')
                            reader.InnerReader.Read();

                        var text = ReadInnerText(reader, null, true, false, true);
                        //Debug.Print(text);
                        fontIndex = null;
                        encoding = _defaultEncoding;

                        if (!string.IsNullOrEmpty(text))
                            stringBuilder.Append(text);

                        break;
                    }

                    case Consts.HtmlBase:
                    {
                        var text = ReadInnerText(reader, null, true, false, true);

                        if (!string.IsNullOrEmpty(text))
                            stringBuilder.Append(text);

                        break;
                    }

                    case Consts.Background:
                    case Consts.Fillcolor:
                    case Consts.Field:
                        ReadInnerText(reader, null, false, true, false);
                        break;

                    default:

				        switch (reader.TokenType)
				        {
                            case RtfTokenType.GroupEnd:
                                htmlState = false;
                                break;

                            case RtfTokenType.Control:
						        if (!htmlState)
						        {
                                    switch (reader.Keyword)
							        {
								        case "'":

                                            if (FontTable != null && fontIndex.HasValue && fontIndex <= FontTable.Count)
                                            {
                                                var font = FontTable[fontIndex.Value];
                                                encoding = font.Encoding ?? _defaultEncoding;
                                            }

                                            // Convert HEX value directly when we have a single byte charset
                                            if (encoding.IsSingleByte)
									        {
									            if (string.IsNullOrEmpty(hexBuffer))
									                hexBuffer = reader.CurrentToken.Hex;

                                                var buff = new[] { byte.Parse(hexBuffer, NumberStyles.HexNumber) };
                                                hexBuffer = string.Empty;
                                                stringBuilder.Append(encoding.GetString(buff));
                                            }
									        else
									        {
									            // If we have a double byte charset like Chinese then store the value and wait for the next HEX value
									            if (hexBuffer == string.Empty)
									                hexBuffer = reader.CurrentToken.Hex;
									            else
									            {
									                // Append the second HEX value and convert it 
									                var buff = new[]
									                {
									                    byte.Parse(hexBuffer, NumberStyles.HexNumber),
									                    byte.Parse(reader.CurrentToken.Hex, NumberStyles.HexNumber)
									                };

									                stringBuilder.Append(encoding.GetString(buff));

									                // Empty the HEX buffer 
									                hexBuffer = string.Empty;
									            }
									        }
									        break;

                                        case "u":
                                            stringBuilder.Append(HttpUtility.UrlDecode("*", _defaultEncoding));
                                            break;
                                    }
                                }
						        break;

					        case RtfTokenType.ExtKeyword:
					        case RtfTokenType.Keyword:
						        if (!htmlState)
						        {
							        switch (reader.Keyword)
							        {
								        case Consts.Par:
									        stringBuilder.Append(Environment.NewLine);
									        break;

								        case Consts.Line:
									        stringBuilder.Append(Environment.NewLine);
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

                                        case Consts.Pntext:
                                            ReadToEndOfGroup(reader);
                                            break;

                                        //case Consts.Fldinst:
                                        //    break;

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

								                if (value >= 0xD800 && value <= 0xDFFF)
								                {
								                    if (!reader.ParsingHighLowSurrogate)
								                    {
								                        reader.ParsingHighLowSurrogate = true;
								                        reader.HighSurrogateValue = value;
								                    }
								                    else
								                    {
								                        var combined = ((reader.HighSurrogateValue - 0xD800) << 10) + (value - 0xDC00) + 0x10000;
								                        stringBuilder.Append("&#" + combined + ";");
								                        reader.ParsingHighLowSurrogate = false;
								                        reader.HighSurrogateValue = null;
								                    }
								                }
								                else
								                {
								                    reader.ParsingHighLowSurrogate = false;
								                    stringBuilder.Append("&#" + value + ";");
								                }
								            }
								            else
								                stringBuilder.Append("&#" + reader.Parameter + ";");
									        break;
							        }
						        }
						        break;

					        case RtfTokenType.Text:
						        if (htmlState == false)
							        stringBuilder.Append(reader.Keyword);
						        break;
				        }
				        break;
		        }
	        }

	        HtmlContent = stringBuilder.ToString();
        }
        #endregion
	}
}