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
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web;

namespace MsgReader.Rtf
{
	/// <summary>
	/// RTF Document
	/// </summary>
	/// <remarks>
	/// This type is the root of RTF Dom tree structure
	/// </remarks>
	internal class DomDocument : DomElement
	{
        #region Fields
        /// <summary>
        /// default font name
        /// </summary>
        private static readonly string DefaultFontName = SystemFonts.DefaultFont.Name;
        // ReSharper disable once NotAccessedField.Local
        private int _listTextFlag;
        private DocumentFormatInfo _paragraphFormat;
        private bool _startContent;
        private int _tokenCount;

        /// <summary>
        /// text encoding of associate font 
        /// </summary>
        private Encoding _associateFontChartset;

        /// <summary>
        /// The default rtf encoding
        /// </summary>
        private Encoding _defaultEncoding = Encoding.Default;

        /// <summary>
        /// Text encoding of current font
        /// </summary>
        private Encoding _fontChartset;
        #endregion

        #region Constructor
        /// <summary>
        /// initialize instance
        /// </summary>
        public DomDocument()
        {
	        DefaultRowHeight = 400;
	        FooterDistance = 720;
	        HeaderDistance = 720;
	        BottomMargin = 1440;
	        RightMargin = 1800;
	        TopMargin = 1440;
	        LeftMargin = 1800;
	        PaperHeight = 15840;
	        PaperWidth = 12240;
	        Info = new DocumentInfo();
	        FontTable = new Table();
	        ChangeTimesNewRoman = false;
	        Generator = null;
	        LeadingChars = null;
	        FollowingChars = null;
	        OwnerDocument = this;
        }
        #endregion

        #region Properties
        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable UnusedAutoPropertyAccessor.Global
        /// <summary>
        /// Following characters
        /// </summary>
        public string FollowingChars { get; set; }

        /// <summary>
        /// Leading characters
        /// </summary>
        public string LeadingChars { get; set; }

        /// <summary>
        /// Text encoding
        /// </summary>
        internal Encoding RuntimeEncoding
        {
	        get
	        {
		        if (_fontChartset != null)
			        return _fontChartset;

		        if (_associateFontChartset != null)
			        return _associateFontChartset;

		        return _defaultEncoding;
	        }
        }

        /// <summary>
        /// Font table
        /// </summary>
        public Table FontTable { get; set; }

        /// <summary>
        /// Document information
        /// </summary>
        public DocumentInfo Info { get; set; }

        /// <summary>
        /// Document generator
        /// </summary>
        public string Generator { get; set; }

        /// <summary>
        /// Format converter
        /// </summary>
        public string FormatConverter { get; set; }
        
        /// <summary>
        /// Paper width,unit twips
        /// </summary>
        public int PaperWidth { get; set; }

        /// <summary>
        /// Paper height,unit twips
        /// </summary>
        public int PaperHeight { get; set; }

        /// <summary>
        /// Left margin,unit twips
        /// </summary>
        public int LeftMargin { get; set; }

        /// <summary>
        /// Top margin,unit twips
        /// </summary>
        public int TopMargin { get; set; }

        /// <summary>
        /// Right margin,unit twips
        /// </summary>
        public int RightMargin { get; set; }

        /// <summary>
        /// Bottom margin,unit twips
        /// </summary>
        public int BottomMargin { get; set; }

        /// <summary>
        /// Landscape
        /// </summary>
        public bool Landscape { get; set; }

        /// <summary>
        /// Header's distance from the top of the page( Twips)
        /// </summary>
        public int HeaderDistance { get; set; }

        /// <summary>
        /// Footer's distance from the bottom of the page( twips)
        /// </summary>
        public int FooterDistance { get; set; }

        /// <summary>
        /// Client area width,unit twips
        /// </summary>
        public int ClientWidth
        {
	        get
	        {
		        if (Landscape)
			        return PaperHeight - LeftMargin - RightMargin;
		        return PaperWidth - LeftMargin - RightMargin;
	        }
        }

        /// <summary>
        /// Convert "Times new roman" to default font when parse rtf content
        /// </summary>
        public bool ChangeTimesNewRoman { get; set; }

        /// <summary>
        /// Default row's height, in twips.
        /// </summary>
        public int DefaultRowHeight { get; set; }

        /// <summary>
        /// HTML content in RTF
        /// </summary>
        public string HtmlContent { get; set; }
        // ReSharper restore MemberCanBePrivate.Global
        // ReSharper restore UnusedAutoPropertyAccessor.Global
        #endregion

        #region Load
        /// <summary>
        /// Load a rtf document from a string in rtf format and parse content
        /// </summary>
        /// <param name="rtfText">text</param>
        public void LoadRtfText(string rtfText)
        {
	        var reader = new StringReader(rtfText);
	        HtmlContent = null;
	        Elements.Clear();
	        _startContent = false;
	        var rtfReader = new Reader(reader);
	        var format = new DocumentFormatInfo();
	        _paragraphFormat = null;
	        Load(rtfReader, format);
        }

        /// <summary>
        /// Parse an RTF element
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="parentFormat"></param>
        private void Load(Reader reader, DocumentFormatInfo parentFormat)
        {
	        if (reader == null)
		        return;

	        var forbitPard = false;
	        DocumentFormatInfo format;
	        if (_paragraphFormat == null)
		        _paragraphFormat = new DocumentFormatInfo();

	        if (parentFormat == null)
		        format = new DocumentFormatInfo();
	        else
	        {
		        format = parentFormat.Clone();
		        format.NativeLevel = parentFormat.NativeLevel + 1;
	        }

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
                            ReadHtmlContent(reader, format);
                            return;

                        #region Read document information
                        case Consts.Listtable:
                            return;

                        case Consts.ListOverride:
                            // Unknow keyword
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Ansi:
                            break;

                        case Consts.Ansicpg:
                            // Read default encoding
                            _defaultEncoding = Encoding.GetEncoding(reader.Parameter);
                            break;

                        case Consts.Fonttbl:
                            // Read font table
                            ReadFontTable(reader);
                            break;

                        case Consts.ListOverrideTable:
                            break;

                        case Consts.FileTable:
                            // Unsupport file list
                            ReadToEndOfGroup(reader);
                            break; // Finish current level

                        case Consts.Colortbl:
                            return; // Finish current level

                        case Consts.StyleSheet:
                            // Unsupport style sheet list
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Generator:
                            // Read document generator
                            Generator = ReadInnerText(reader, true);
                            break;

                        case Consts.Info:
                            // Read document information
                            ReadDocumentInfo(reader);
                            return;

                        case Consts.Headery:
                            if (reader.HasParam)
                                HeaderDistance = reader.Parameter;
                            break;

                        case Consts.Footery:
                            if (reader.HasParam)
                                FooterDistance = reader.Parameter;
                            break;

                        case Consts.Header:
                            break;

                        case Consts.Headerl:
                            break;

                        case Consts.Headerr:
                            break;

                        case Consts.Headerf:
                            break;

                        case Consts.Footer:
                            break;

                        case Consts.Footerl:
                            break;

                        case Consts.Footerr:
                            break;

                        case Consts.Footerf:
                            break;

                        case Consts.Xmlns:
                            // Unsupport xml namespace
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Nonesttables:
                            // I support nest table , then ignore this keyword
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Xmlopen:
                            // Unsupport xmlopen keyword
                            break;

                        case Consts.Revtbl:
                            //ReadToEndGround(reader);
                            break;
                        #endregion

                        #region Read document information
                        case Consts.Paperw:
                            // Read paper width
                            PaperWidth = reader.Parameter;
                            break;

                        case Consts.Paperh:
                            // Read paper height
                            PaperHeight = reader.Parameter;
                            break;

                        case Consts.Margl:
                            // Read left margin
                            LeftMargin = reader.Parameter;
                            break;

                        case Consts.Margr:
                            // Read right margin
                            RightMargin = reader.Parameter;
                            break;

                        case Consts.Margb:
                            // Read bottom margin
                            BottomMargin = reader.Parameter;
                            break;

                        case Consts.Margt:
                            // Read top margin 
                            TopMargin = reader.Parameter;
                            break;

                        case Consts.Landscape:
                            // Set landscape
                            Landscape = true;
                            break;

                        case Consts.Fchars:
                            FollowingChars = ReadInnerText(reader, true);
                            break;

                        case Consts.Lchars:
                            LeadingChars = ReadInnerText(reader, true);
                            break;

                        case "pnseclvl":
                            // Ignore this keyword
                            ReadToEndOfGroup(reader);
                            break;
                        #endregion

                        #region Read paragraph format
                        case Consts.Pard:
                            _startContent = true;
                            if (forbitPard)
                                continue;

                            // Clear paragraph format
                            _paragraphFormat.ResetParagraph();
                            // Format.ResetParagraph();
                            break;

                        case Consts.Par:
                            break;

                        case Consts.Page:
                            break;

                        case Consts.Pagebb:
                            _startContent = true;
                            _paragraphFormat.PageBreak = true;
                            break;

                        case Consts.Ql:
                            // Left alignment
                            _startContent = true;
                            _paragraphFormat.Align = RtfAlignment.Left;
                            break;

                        case Consts.Qc:
                            // Center alignment
                            _startContent = true;
                            _paragraphFormat.Align = RtfAlignment.Center;
                            break;

                        case Consts.Qr:
                            // Right alignment
                            _startContent = true;
                            _paragraphFormat.Align = RtfAlignment.Right;
                            break;

                        case Consts.Qj:
                            // Jusitify alignment
                            _startContent = true;
                            _paragraphFormat.Align = RtfAlignment.Justify;
                            break;

                        case Consts.Sl:
                            // Line spacing
                            _startContent = true;
                            if (reader.Parameter >= 0)
                                _paragraphFormat.LineSpacing = reader.Parameter;
                            break;

                        case Consts.Slmult:
                            _startContent = true;
                            _paragraphFormat.MultipleLineSpacing = (reader.Parameter == 1);
                            break;

                        case Consts.Sb:
                            // Spacing before paragraph
                            _startContent = true;
                            _paragraphFormat.SpacingBefore = reader.Parameter;
                            break;

                        case Consts.Sa:
                            // Spacing after paragraph
                            _startContent = true;
                            _paragraphFormat.SpacingAfter = reader.Parameter;
                            break;

                        case Consts.Fi:
                            // Indent first line
                            _startContent = true;
                            _paragraphFormat.ParagraphFirstLineIndent = reader.Parameter;
                            break;

                        case Consts.Brdrw:
                            _startContent = true;
                            if (reader.HasParam)
                                _paragraphFormat.BorderWidth = reader.Parameter;
                            break;

                        case Consts.Pn:
                            _startContent = true;
                            _paragraphFormat.ListId = -1;
                            break;

                        case Consts.Pntext:
                            break;

                        case Consts.Pntxtb:
                            break;

                        case Consts.Pntxta:
                            break;

                        case Consts.Pnlvlbody:
                            _startContent = true;
                            break;

                        case Consts.Pnlvlblt:
                            _startContent = true;
                            break;

                        case Consts.Listtext:
                            _startContent = true;
                            var text = ReadInnerText(reader, true);
                            if (text != null)
                            {
                                text = text.Trim();
                                _listTextFlag = text.StartsWith("l") ? 1 : 2;
                            }

                            break;

                        case Consts.Ls:
                            _startContent = true;
                            _paragraphFormat.ListId = reader.Parameter;
                            _listTextFlag = 0;
                            break;

                        case Consts.Li:
                            _startContent = true;
                            if (reader.HasParam)
                                _paragraphFormat.LeftIndent = reader.Parameter;
                            break;

                        case Consts.Line:
                            break;
                        #endregion

                        #region Read text format
                        case Consts.Insrsid:
                            break;

                        case Consts.Plain:
                            // Clear text format
                            _startContent = true;
                            format.ResetText();
                            break;

                        case Consts.F:
                            // Font name
                            _startContent = true;
                            if (format.ReadText)
                            {
                                var fontName = FontTable.GetFontName(reader.Parameter);

                                fontName = fontName?.Trim();

                                if (string.IsNullOrEmpty(fontName))
                                    fontName = DefaultFontName;

                                if (ChangeTimesNewRoman)
                                {
                                    if (fontName == "Times New Roman")
                                        fontName = DefaultFontName;
                                }

                                format.FontName = fontName;
                            }

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

                        case Consts.Fs:
                            // Font size
                            _startContent = true;
                            if (format.ReadText)
                            {
                                if (reader.HasParam)
                                    format.FontSize = reader.Parameter / 2.0f;
                            }

                            break;

                        case Consts.Cf:
                            break;

                        case Consts.Cb:
                        case Consts.Chcbpat:
                            break;

                        case Consts.B:
                            break;

                        case Consts.V:
                            break;

                        case Consts.Highlight:
                            break;

                        case Consts.I:
                        case Consts.Ul:
                        case Consts.Strike:
                        case Consts.Sub:
                        case Consts.Super:
                        case Consts.Nosupersub:
                        case Consts.Brdrb:
                        case Consts.Brdrl:
                        case Consts.Brdrr:
                        case Consts.Brdrt:
                        case Consts.Brdrcf:
                            break;

                        case Consts.Brdrs:
                            _startContent = true;
                            _paragraphFormat.BorderThickness = false;
                            format.BorderThickness = false;
                            break;

                        case Consts.Brdrth:
                            _startContent = true;
                            _paragraphFormat.BorderThickness = true;
                            format.BorderThickness = true;
                            break;

                        case Consts.Brdrdot:
                            _startContent = true;
                            _paragraphFormat.BorderStyle = DashStyle.Dot;
                            format.BorderStyle = DashStyle.Dot;
                            break;

                        case Consts.Brdrdash:
                            _startContent = true;
                            _paragraphFormat.BorderStyle = DashStyle.Dash;
                            format.BorderStyle = DashStyle.Dash;
                            break;

                        case Consts.Brdrdashd:
                            _startContent = true;
                            _paragraphFormat.BorderStyle = DashStyle.DashDot;
                            format.BorderStyle = DashStyle.DashDot;
                            break;

                        case Consts.Brdrdashdd:
                            _startContent = true;
                            _paragraphFormat.BorderStyle = DashStyle.DashDotDot;
                            format.BorderStyle = DashStyle.DashDotDot;
                            break;

                        case Consts.Brdrnil:
                            _startContent = true;
                            _paragraphFormat.LeftBorder = false;
                            _paragraphFormat.TopBorder = false;
                            _paragraphFormat.RightBorder = false;
                            _paragraphFormat.BottomBorder = false;

                            format.LeftBorder = false;
                            format.TopBorder = false;
                            format.RightBorder = false;
                            format.BottomBorder = false;
                            break;

                        case Consts.Brsp:
                            _startContent = true;
                            if (reader.HasParam)
                                _paragraphFormat.BorderSpacing = reader.Parameter;
                            break;

                        case Consts.Chbrdr:
                            _startContent = true;
                            format.LeftBorder = true;
                            format.TopBorder = true;
                            format.RightBorder = true;
                            format.BottomBorder = true;
                            break;

                        case Consts.Bkmkstart:
                            break;

                        case Consts.Bkmkend:
                            forbitPard = true;
                            format.ReadText = false;
                            break;

                        case Consts.Field:
                            return; // finish current level
                        //break;
                        #endregion

                        case Consts.Object:
                            return; // finish current level

                        #region Read image
                        case Consts.Shppict:
                            // Continue the following token
                            break;

                        case Consts.Nonshppict:
                            // unsupport keyword
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Pict:
                            break;

                        case Consts.Picscalex:
                            break;

                        case Consts.Picscaley:
                            break;

                        case Consts.Picwgoal:
                            break;

                        case Consts.Pichgoal:
                            break;

                        case Consts.Blipuid:
                            break;

                        case Consts.Emfblip:
                            break;

                        case Consts.Pngblip:
                            break;

                        case Consts.Jpegblip:
                            break;

                        case Consts.Macpict:
                            break;

                        case Consts.Pmmetafile:
                            break;

                        case Consts.Wmetafile:
                            break;

                        case Consts.Dibitmap:
                            break;

                        case Consts.Wbitmap:
                            break;
                        #endregion

                        #region Read shape
                        case Consts.Sp:
                            break;

                        case Consts.Shptxt:
                            // handle following token
                            break;

                        case Consts.Shprslt:
                            // ignore this level
                            ReadToEndOfGroup(reader);
                            break;

                        case Consts.Shp:
                            break;

                        case Consts.Shpleft:
                            break;

                        case Consts.Shptop:
                            break;

                        case Consts.Shpright:
                            break;

                        case Consts.Shpbottom:
                            break;

                        case Consts.Shplid:
                            break;

                        case Consts.Shpz:
                            break;

                        case Consts.Shpgrp:
                            break;

                        case Consts.Shpinst:
                            break;
                        #endregion

                        #region Read table
                        case Consts.Intbl:
                        case Consts.Trowd:
                        case Consts.Itap:
                            break;

                        case Consts.Nesttableprops:
                            break;

                        case Consts.Row:
                            break;

                        case Consts.Nestrow:
                            break;

                        case Consts.Trrh:
                        case Consts.Trautofit:
                        case Consts.Irowband:
                        case Consts.Trhdr:
                        case Consts.Trkeep:
                        case Consts.Trkeepfollow:
                        case Consts.Trleft:
                        case Consts.Trqc:
                        case Consts.Trql:
                        case Consts.Trqr:
                        case Consts.Trcbpat:
                        case Consts.Trcfpat:
                        case Consts.Trpat:
                        case Consts.Trshdng:
                        case Consts.TrwWidth:
                        case Consts.TrwWidthA:
                        case Consts.Irow:
                        case Consts.Trpaddb:
                        case Consts.Trpaddl:
                        case Consts.Trpaddr:
                        case Consts.Trpaddt:
                        case Consts.Trpaddfb:
                        case Consts.Trpaddfl:
                        case Consts.Trpaddfr:
                        case Consts.Trpaddft:
                        case Consts.Lastrow:
                            break;

                        case Consts.Clvmgf:
                        case Consts.Clvmrg:
                        case Consts.Cellx:
                        case Consts.Clvertalt:
                        case Consts.Clvertalc:
                        case Consts.Clvertalb:
                        case Consts.ClNoWrap:
                        case Consts.Clcbpat:
                        case Consts.Clcfpat:
                        case Consts.Clpadl:
                        case Consts.Clpadt:
                        case Consts.Clpadr:
                        case Consts.Clpadb:
                        case Consts.Clbrdrl:
                        case Consts.Clbrdrt:
                        case Consts.Clbrdrr:
                        case Consts.Clbrdrb:
                        case Consts.Brdrtbl:
                        case Consts.Brdrnone:
                            break;

                        case Consts.Cell:
                            break;

                        case Consts.Nestcell:
                            break;
                        #endregion

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
        #endregion

        #region GetLastElements
        /// <summary>
        /// Get the last element
        /// </summary>
        /// <param name="checkLockState"></param>
        /// <returns></returns>
        private DomElement[] GetLastElements(bool checkLockState)
        {
	        var result = new List<DomElement>();
	        DomElement element = this;
	        while (element != null)
	        {
		        if (checkLockState)
		        {
			        if (element.Locked)
			        {
				        break;
			        }
		        }
		        result.Add(element);
		        element = element.Elements.LastElement;
	        }
	        if (checkLockState)
	        {
		        for (var count = result.Count - 1; count >= 0; count--)
		        {
			        if (result[count].Locked)
			        {
				        result.RemoveAt(count);
			        }
		        }
	        }
	        return result.ToArray();
        }
        #endregion

        #region HexToBytes
        /// <summary>
        /// Convert a hex string to a byte array
        /// </summary>
        /// <param name="hex">hex string</param>
        /// <returns>byte array</returns>
        private byte[] HexToBytes(string hex)
        {
	        const string chars = "0123456789abcdef";

	        var value = 0;
	        var charCount = 0;
	        var buffer = new ByteBuffer();
	        for (var count = 0; count < hex.Length; count++)
	        {
		        var c = hex[count];
		        c = char.ToLower(c);
		        var index = chars.IndexOf(c);
		        if (index >= 0)
		        {
			        charCount++;
			        value = value * 16 + index;
			        if (charCount > 0 && (charCount % 2) == 0)
			        {
				        buffer.Add((byte)value);
				        value = 0;
			        }
		        }
	        }
	        return buffer.ToArray();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
	        return "RTFDocument:" + Info.Title;
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
        /// <param name="format"></param>
        private void ReadHtmlContent(Reader reader, DocumentFormatInfo format)
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