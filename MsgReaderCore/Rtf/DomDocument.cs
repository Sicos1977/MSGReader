using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

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
		#region Events
		/// <summary>
		/// Progress event
		/// </summary>
		public event ProgressEventHandler Progress;
		#endregion

		#region Fields
		/// <summary>
		/// default font name
		/// </summary>
		private static readonly string DefaultFontName = Control.DefaultFont.Name;
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
			ListOverrideTable = new ListOverrideTable();
			ListTable = new ListTable();
			ColorTable = new ColorTable();
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
		/// Color table
		/// </summary>
		public ColorTable ColorTable { get; set; }

		public ListTable ListTable { get; set; }

		public ListOverrideTable ListOverrideTable { get; set; }

		/// <summary>
		/// Document information
		/// </summary>
		public DocumentInfo Info { get; set; }

		/// <summary>
		/// Document generator
		/// </summary>
		public string Generator { get; set; }

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

		#region OnProgress
		/// <summary>
		/// Raise progress event
		/// </summary>
		/// <param name="max">Progress max value</param>
		/// <param name="value">Progress value</param>
		/// <param name="message">Progress message</param>
		/// <returns>User cancel</returns>
		private void OnProgress(int max, int value, string message)
		{
			if (Progress != null)
			{
				var args = new ProgressEventArgs(max, value, message);
				Progress(this, args);
			}
		}
		#endregion

		#region Load
		/// <summary>
		/// Load a rtf file and parse
		/// </summary>
		/// <param name="fileName">file name</param>
		public void Load(string fileName)
		{
			using (var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
				Load(fileStream);
		}

		/// <summary>
		/// Load a rtf document from a stream and parse content
		/// </summary>
		/// <param name="stream">stream</param>
		public void Load(Stream stream)
		{
			HtmlContent = null;
			Elements.Clear();
			_startContent = false;
			var reader = new Reader(stream);
			var format = new DocumentFormatInfo();
			_paragraphFormat = null;
			Load(reader, format);
			// Combination table rows to table
			CombineTable(this);
			FixElements(this);
		}

		/// <summary>
		/// Load a rtf document from a text reader and parse content
		/// </summary>
		/// <param name="reader">text reader</param>
		public void Load(TextReader reader)
		{
			HtmlContent = null;
			Elements.Clear();
			_startContent = false;
			var r = new Reader(reader);
			var format = new DocumentFormatInfo();
			_paragraphFormat = null;
			Load(r, format);
			// combination table rows to table
			CombineTable(this);
			FixElements(this);
		}

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
			CombineTable(this);
			FixElements(this);
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

			var textContainer = new TextContainer(this);
			var levelBack = reader.Level;

			while (reader.ReadToken() != null)
			{
				if (reader.TokenCount - _tokenCount > 100)
				{
					_tokenCount = reader.TokenCount;
					OnProgress(reader.ContentLength, reader.ContentPosition, null);
				}
				if (_startContent)
				{
					if (textContainer.Accept(reader.CurrentToken, reader))
					{
						textContainer.Level = reader.Level;
						continue;
					}
					if (textContainer.HasContent)
					{
						if (ApplyText(textContainer, reader, format))
							break;
					}
				}

				if (reader.TokenType == RtfTokenType.GroupEnd)
				{
					var elements = GetLastElements(true);
					for (var count = 0; count < elements.Length; count++)
					{
						var element = elements[count];
						if (element.NativeLevel >= 0 && element.NativeLevel > reader.Level)
						{
							for (var count2 = count; count2 < elements.Length; count2++)
								elements[count2].Locked = true;
							break;
						}
					}

					break;
				}

				if (reader.Level < levelBack)
					break;

				if (reader.TokenType == RtfTokenType.GroupStart)
				{
					Load(reader, format);
					if (reader.Level < levelBack)
						break;
				}

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

						#region Read document information
						case Consts.Listtable:
							ReadListTable(reader);
							return;

						case Consts.ListOverride:
							// Unknow keyword
							ReadToEndGround(reader);
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
							ReadListOverrideTable(reader);
							break;

						case Consts.FileTable:
							// Unsupport file list
							ReadToEndGround(reader);
							break; // Finish current level

						case Consts.Colortbl:
							// Read color table
							ReadColorTable(reader);
							return; // Finish current level

						case Consts.StyleSheet:
							// Unsupport style sheet list
							ReadToEndGround(reader);
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
							// Analyse header
							var header = new DomHeader { Style = RtfHeaderFooterStyle.AllPages };
							AppendChild(header);
							Load(reader, parentFormat);
							header.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Headerl:
							// Analyse header
							var header1 = new DomHeader { Style = RtfHeaderFooterStyle.LeftPages };
							AppendChild(header1);
							Load(reader, parentFormat);
							header1.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Headerr:
							// Analyse header
							var headerr = new DomHeader { Style = RtfHeaderFooterStyle.RightPages };
							AppendChild(headerr);
							Load(reader, parentFormat);
							headerr.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Headerf:
							// Analyse header
							var headerf = new DomHeader { Style = RtfHeaderFooterStyle.FirstPage };
							AppendChild(headerf);
							Load(reader, parentFormat);
							headerf.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Footer:
							// Analyse footer
							var footer = new DomFooter { Style = RtfHeaderFooterStyle.FirstPage };
							AppendChild(footer);
							Load(reader, parentFormat);
							footer.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Footerl:
							// analyse footer
							var footerl = new DomFooter { Style = RtfHeaderFooterStyle.LeftPages };
							AppendChild(footerl);
							Load(reader, parentFormat);
							footerl.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Footerr:
							// Analyse footer
							var footerr = new DomFooter { Style = RtfHeaderFooterStyle.RightPages };
							AppendChild(footerr);
							Load(reader, parentFormat);
							footerr.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Footerf:
							// analyse footer
							var footerf = new DomFooter { Style = RtfHeaderFooterStyle.FirstPage };
							AppendChild(footerf);
							Load(reader, parentFormat);
							footerf.Locked = true;
							_paragraphFormat = new DocumentFormatInfo();
							break;

						case Consts.Xmlns:
							// Unsupport xml namespace
							ReadToEndGround(reader);
							break;

						case Consts.Nonesttables:
							// I support nest table , then ignore this keyword
							ReadToEndGround(reader);
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
							ReadToEndGround(reader);
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
							_startContent = true;
							// New paragraph
							if (GetLastElement(typeof(DomParagraph)) == null)
							{
								var paragraph = new DomParagraph { Format = _paragraphFormat };
								_paragraphFormat = _paragraphFormat.Clone();
								AddContentElement(paragraph);
								paragraph.Locked = true;
							}
							else
							{
								CompleteParagraph();
								var p = new DomParagraph { Format = _paragraphFormat };
								AddContentElement(p);
							}
							_startContent = true;
							break;

						case Consts.Page:
							_startContent = true;
							CompleteParagraph();
							AddContentElement(new DomPageBreak());
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
							_startContent = true;
							if (format.ReadText)
							{
								var line = new DomLineBreak();
								line.NativeLevel = reader.Level;
								AddContentElement(line);
							}
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
								if (fontName != null)
									fontName = fontName.Trim();
								if (string.IsNullOrEmpty(fontName))
									fontName = DefaultFontName;

								if (ChangeTimesNewRoman)
								{
									if (fontName == "Times New Roman")
										fontName = DefaultFontName;
								}
								format.FontName = fontName;
							}
							_fontChartset = FontTable[reader.Parameter].Encoding;
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
							// Font color
							_startContent = true;
							if (format.ReadText)
							{
								if (reader.HasParam)
									format.TextColor = ColorTable.GetColor(reader.Parameter, Color.Black);
							}
							break;

						case Consts.Cb:
						case Consts.Chcbpat:
							// Background color
							_startContent = true;
							if (format.ReadText)
							{
								if (reader.HasParam)
								{
									format.BackColor = ColorTable.GetColor(reader.Parameter, Color.Empty);
								}
							}
							break;

						case Consts.B:
							// Bold
							_startContent = true;
							if (format.ReadText)
								format.Bold = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.V:
							// Hidden text
							_startContent = true;
							if (format.ReadText)
							{
								if (reader.HasParam && reader.Parameter == 0)
									format.Hidden = false;
								else
									format.Hidden = true;
							}
							break;

						case Consts.Highlight:
							// Highlight content
							_startContent = true;
							if (format.ReadText)
							{
								if (reader.HasParam)
									format.BackColor = ColorTable.GetColor(
										reader.Parameter,
										Color.Empty);
							}
							break;

						case Consts.I:
							// Italic
							_startContent = true;
							if (format.ReadText)
								format.Italic = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.Ul:
							// Under line
							_startContent = true;
							if (format.ReadText)
								format.Underline = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.Strike:
							// Strikeout
							_startContent = true;
							if (format.ReadText)
								format.Strikeout = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.Sub:
							// Subscript
							_startContent = true;
							if (format.ReadText)
								format.Subscript = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.Super:
							// superscript
							_startContent = true;
							if (format.ReadText)
								format.Superscript = (reader.HasParam == false || reader.Parameter != 0);
							break;

						case Consts.Nosupersub:
							// nosupersub
							_startContent = true;
							format.Subscript = false;
							format.Superscript = false;
							break;

						case Consts.Brdrb:
							_startContent = true;
							//format.ParagraphBorder.Bottom = true;
							_paragraphFormat.BottomBorder = true;
							break;

						case Consts.Brdrl:
							_startContent = true;
							//format.ParagraphBorder.Left = true ;
							_paragraphFormat.LeftBorder = true;
							break;

						case Consts.Brdrr:
							_startContent = true;
							//format.ParagraphBorder.Right = true ;
							_paragraphFormat.RightBorder = true;
							break;

						case Consts.Brdrt:
							_startContent = true;
							//format.ParagraphBorder.Top = true;
							_paragraphFormat.BottomBorder = true;
							break;

						case Consts.Brdrcf:
							_startContent = true;
							var element = GetLastElement(typeof(DomTableRow), false);
							if (element is DomTableRow)
							{
								// Reading a table row
								var row = (DomTableRow)element;
								if (row.CellSettings.Count > 0)
								{
									var style = (AttributeList)row.CellSettings[row.CellSettings.Count - 1];
									style.Add(reader.Keyword, reader.Parameter);
								}
							}
							else
							{
								_paragraphFormat.BorderColor = ColorTable.GetColor(reader.Parameter, Color.Black);
								format.BorderColor = format.BorderColor;
							}
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
							// Book mark
							_startContent = true;
							if (format.ReadText && _startContent)
							{
								var bk = new DomBookmark();
								bk.Name = ReadInnerText(reader, true);
								bk.Locked = true;
								AddContentElement(bk);
							}
							break;

						case Consts.Bkmkend:
							forbitPard = true;
							format.ReadText = false;
							break;

						case Consts.Field:
							// Field
							_startContent = true;
							ReadDomField(reader, format);
							return; // finish current level
									//break;
						#endregion

						#region Read object
						case Consts.Object:
							{
								// object
								_startContent = true;
								ReadDomObject(reader, format);
								return; // finish current level
							}
						#endregion

						#region Read image
						case Consts.Shppict:
							// Continue the following token
							break;

						case Consts.Nonshppict:
							// unsupport keyword
							ReadToEndGround(reader);
							break;

						case Consts.Pict:
							{
								// Read image data
								//ReadDomImage(reader, format);
								_startContent = true;
								var image = new DomImage();
								image.NativeLevel = reader.Level;
								AddContentElement(image);
								break;
							}

						case Consts.Picscalex:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.ScaleX = reader.Parameter;
								break;
							}

						case Consts.Picscaley:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.ScaleY = reader.Parameter;
								break;
							}

						case Consts.Picwgoal:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.DesiredWidth = reader.Parameter;
								break;
							}

						case Consts.Pichgoal:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.DesiredHeight = reader.Parameter;
								break;
							}

						case Consts.Blipuid:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.Id = ReadInnerText(reader, true);
								break;
							}

						case Consts.Emfblip:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Emfblip;
								break;
							}

						case Consts.Pngblip:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Pngblip;
								break;
							}

						case Consts.Jpegblip:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Jpegblip;
								break;
							}

						case Consts.Macpict:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Macpict;
								break;
							}

						case Consts.Pmmetafile:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Pmmetafile;
								break;
							}

						case Consts.Wmetafile:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Wmetafile;
								break;
							}

						case Consts.Dibitmap:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Dibitmap;
								break;
							}

						case Consts.Wbitmap:
							{
								var image = (DomImage)GetLastElement(typeof(DomImage));
								if (image != null)
									image.PicType = RtfPictureType.Wbitmap;
								break;
							}
						#endregion

						#region Read shape
						case Consts.Sp:
							{
								// Begin read shape property
								var level = 0;
								string vName = null;
								string vValue = null;
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
									else switch (reader.Keyword)
										{
											case Consts.Sn:
												vName = ReadInnerText(reader, true);
												break;
											case Consts.Sv:
												vValue = ReadInnerText(reader, true);
												break;
										}
								}

								var shape = (DomShape)GetLastElement(typeof(DomShape));

								if (shape != null)
									shape.ExtAttrbutes[vName] = vValue;
								else
								{
									var g = (DomShapeGroup)GetLastElement(typeof(DomShapeGroup));
									if (g != null)
									{
										g.ExtAttrbutes[vName] = vValue;
									}
								}
								break;
							}

						case Consts.Shptxt:
							// handle following token
							break;

						case Consts.Shprslt:
							// ignore this level
							ReadToEndGround(reader);
							break;

						case Consts.Shp:
							{
								_startContent = true;
								var shape = new DomShape();
								shape.NativeLevel = reader.Level;
								AddContentElement(shape);
								break;
							}
						case Consts.Shpleft:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.Left = reader.Parameter;
								break;
							}

						case Consts.Shptop:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.Top = reader.Parameter;
								break;
							}

						case Consts.Shpright:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.Width = reader.Parameter - shape.Left;
								break;
							}

						case Consts.Shpbottom:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.Height = reader.Parameter - shape.Top;
								break;
							}

						case Consts.Shplid:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.Id = reader.Parameter;
								break;
							}

						case Consts.Shpz:
							{
								var shape = (DomShape)GetLastElement(typeof(DomShape));
								if (shape != null)
									shape.ZIndex = reader.Parameter;
								break;
							}

						case Consts.Shpgrp:
							{
								var group = new DomShapeGroup();
								group.NativeLevel = reader.Level;
								AddContentElement(group);
								break;
							}

						case Consts.Shpinst:
							break;
						#endregion

						#region Read table
						case Consts.Intbl:
						case Consts.Trowd:
						case Consts.Itap:
							{
								// These keyword said than current paragraph is a table row
								_startContent = true;
								var es = GetLastElements(true);
								DomElement lastUnlockElement = null;
								DomElement lastTableElement = null;
								for (var count = es.Length - 1; count >= 0; count--)
								{
									var e = es[count];
									if (e.Locked == false)
									{
										if (lastUnlockElement == null && !(e is DomParagraph))
											lastUnlockElement = e;
										if (e is DomTableRow || e is DomTableCell)
											lastTableElement = e;
										break;
									}
								}

								if (reader.Keyword == Consts.Intbl)
								{
									if (lastTableElement == null)
									{
										// if can not find unlocked row 
										// then new row
										var row = new DomTableRow { NativeLevel = reader.Level };
										if (lastUnlockElement != null)
											lastUnlockElement.AppendChild(row);
									}
								}
								else if (reader.Keyword == Consts.Trowd)
								{
									// clear row format
									DomTableRow row;
									if (lastTableElement == null)
									{
										row = new DomTableRow { NativeLevel = reader.Level };
										if (lastUnlockElement != null)
											lastUnlockElement.AppendChild(row);
									}
									else
									{
										row = lastTableElement as DomTableRow ?? (DomTableRow)lastTableElement.Parent;
									}
									row.Attributes.Clear();
									row.CellSettings.Clear();
									_paragraphFormat.ResetParagraph();
								}
								else if (reader.Keyword == Consts.Itap)
								{
									// set nested level

									if (reader.Parameter == 0)
									{
										// is the 0 level , belong to document , not to a table
										//foreach (RTFDomElement element in es)
										//{
										//    if (element is RTFDomTableRow || element is RTFDomTableCell)
										//    {
										//        element.Locked = true;
										//    }
										//}
									}
									else
									{
										// in a row
										DomTableRow row;
										if (lastTableElement == null)
										{
											row = new DomTableRow { NativeLevel = reader.Level };
											if (lastUnlockElement != null)
												lastUnlockElement.AppendChild(row);
										}
										else
											row = lastTableElement as DomTableRow ?? (DomTableRow)lastTableElement.Parent;
										if (reader.Parameter == row.Level)
										{
										}
										else if (reader.Parameter > row.Level)
										{
											// nested row
											var newRow = new DomTableRow { Level = reader.Parameter };
											var parentCell = (DomTableCell)GetLastElement(typeof(DomTableCell), false);
											if (parentCell == null)
												AddContentElement(newRow);
											else
												parentCell.AppendChild(newRow);
										}
										else if (reader.Parameter < row.Level)
										{
											// exit nested row
										}
									}
								}
								break;
							}

						case Consts.Nesttableprops:
							// ignore
							break;

						case Consts.Row:
							{
								// finish read row
								_startContent = true;
								var es = GetLastElements(true);
								for (var count = es.Length - 1; count >= 0; count--)
								{
									es[count].Locked = true;
									if (es[count] is DomTableRow)
										break;
								}
								break;
							}

						case Consts.Nestrow:
							{
								// finish nested row
								_startContent = true;
								var es = GetLastElements(true);
								for (var count = es.Length - 1; count >= 0; count--)
								{
									es[count].Locked = true;
									if (es[count] is DomTableRow)
										break;
								}
								break;
							}

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
							{
								// meet row control word , not parse at first , just save it 
								_startContent = true;
								var row = (DomTableRow)GetLastElement(typeof(DomTableRow), false);
								if (row != null)
								{
									row.Attributes.Add(reader.Keyword, reader.Parameter);
								}
								break;
							}

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
							{
								// Meet cell control word , no parse at first , just save it
								var row = (DomTableRow)GetLastElement(typeof(DomTableRow), false);
								if (row == null) break;
								_startContent = true;
								AttributeList style = null;
								if (row.CellSettings.Count > 0)
								{
									style = (AttributeList)row.CellSettings[row.CellSettings.Count - 1];
									if (style.Contains(Consts.Cellx))
									{
										// if find repeat control word , then can consider this control word
										// belong to the next cell . userly cellx is the last control word of 
										// a cell , when meet cellx , the current cell defind is finished.
										style = new AttributeList();
										row.CellSettings.Add(style);
									}
								}
								if (style == null)
								{
									style = new AttributeList();
									row.CellSettings.Add(style);
								}
								style.Add(reader.Keyword, reader.Parameter);
								break;
							}

						case Consts.Cell:
							{
								// finish cell content
								_startContent = true;
								AddContentElement(null);
								CompleteParagraph();
								_paragraphFormat.Reset();
								format.Reset();
								var es = GetLastElements(true);
								for (var count = es.Length - 1; count >= 0; count--)
								{
									if (es[count].Locked == false)
									{
										es[count].Locked = true;
										if (es[count] is DomTableCell)
											break;
									}
								}
								break;
							}

						case Consts.Nestcell:
							{
								// finish nested cell content
								_startContent = true;
								AddContentElement(null);
								CompleteParagraph();
								var es = GetLastElements(false);
								for (var count = es.Length - 1; count >= 0; count--)
								{
									es[count].Locked = true;
									if (es[count] is DomTableCell)
									{
										((DomTableCell)es[count]).Format = format;
										break;
									}
								}
								break;
							}
						#endregion

						default:
							// Unsupport keyword
							if (reader.TokenType == RtfTokenType.ExtKeyword && reader.FirstTokenInGroup)
							{
								// if meet unsupport extern keyword , and this token is the first token in 
								// current group , then ingore whole group.
								ReadToEndGround(reader);
							}
							break;
					}
				}
			}

			if (textContainer.HasContent)
				ApplyText(textContainer, reader, format);
		}
		#endregion

		#region FixForParagraphs
		/// <summary>
		/// Fixes invalid paragraphs
		/// </summary>
		// ReSharper disable UnusedMember.Global
		public void FixForParagraphs(DomElement parentElement)
		// ReSharper restore UnusedMember.Global
		{
			DomParagraph lastParagraph = null;
			var list = new DomElementList();
			foreach (DomElement element in parentElement.Elements)
			{
				if (element is DomHeader
					|| element is DomFooter)
				{
					FixForParagraphs(element);
					lastParagraph = null;
					list.Add(element);
					continue;
				}

				if (element is DomParagraph
					|| element is DomTableRow
					|| element is DomTable
					|| element is DomTableCell)
				{
					lastParagraph = null;
					list.Add(element);
					continue;
				}

				if (lastParagraph == null)
				{
					lastParagraph = new DomParagraph();
					list.Add(lastParagraph);
					if (element is DomText)
						lastParagraph.Format = ((DomText)element).Format.Clone();
				}

				lastParagraph.Elements.Add(element);
			}

			parentElement.Elements.Clear();

			foreach (DomElement element in list)
				parentElement.Elements.Add(element);
		}
		#endregion

		#region FixElements
		/// <summary>
		/// Fixes invalid dom elements
		/// </summary>
		/// <param name="parentElement"></param>
		private void FixElements(DomElement parentElement)
		{
			// combin text element , decrease number of RTFDomText instance
			var result = new ArrayList();
			foreach (DomElement element in parentElement.Elements)
			{
				if (element is DomParagraph)
				{
					var p = (DomParagraph)element;
					if (p.Format.PageBreak)
					{
						p.Format.PageBreak = false;
						result.Add(new DomPageBreak());
					}
				}

				if (element is DomText)
				{
					if (result.Count > 0 && result[result.Count - 1] is DomText)
					{
						var lastText = (DomText)result[result.Count - 1];
						var txt = (DomText)element;
						if (lastText.Text.Length == 0 || txt.Text.Length == 0)
						{
							if (lastText.Text.Length == 0)
							{
								// close text format
								lastText.Format = txt.Format.Clone();
							}
							lastText.Text = lastText.Text + txt.Text;
						}
						else
						{
							if (lastText.Format.EqualsSettings(txt.Format))
							{
								lastText.Text = lastText.Text + txt.Text;
							}
							else
							{
								result.Add(txt);
							}
						}
					}
					else
						result.Add(element);
				}
				else
					result.Add(element);
			}

			parentElement.Elements.Clear();
			parentElement.Locked = false;

			foreach (DomElement element in result)
				parentElement.AppendChild(element);

			foreach (var element in parentElement.Elements.ToArray())
			{
				if (element is DomTable)
					UpdateTableCells((DomTable)element, true);
			}

			// Recursive
			foreach (DomElement element in parentElement.Elements)
				FixElements(element);
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

		private DomElement GetLastElement(Type elementType)
		{
			var elements = GetLastElements(true);
			for (var count = elements.Length - 1; count >= 0; count--)
			{
				if (elementType.IsInstanceOfType(elements[count]))
					return elements[count];
			}
			return null;
		}

		private DomElement GetLastElement(Type elementType, bool lockStatus)
		{
			var elements = GetLastElements(true);
			for (var count = elements.Length - 1; count >= 0; count--)
			{
				if (elementType.IsInstanceOfType(elements[count]))
				{
					if (elements[count].Locked == lockStatus)
					{
						return elements[count];
					}
				}
			}
			return null;
		}

		private DomElement GetLastElement()
		{
			var elements = GetLastElements(true);
			if (elements.Length == 0) return null;
			return elements[elements.Length - 1];
		}
		#endregion

		#region CompleteParagraph
		/// <summary>
		/// Complete the paragraph
		/// </summary>
		private void CompleteParagraph()
		{
			var lastElement = GetLastElement();
			while (lastElement != null)
			{
				// ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
				if (lastElement is DomParagraph)
				{
					var p = (DomParagraph)lastElement;
					p.Locked = true;
					if (_paragraphFormat != null)
					{
						p.Format = _paragraphFormat;
						_paragraphFormat = _paragraphFormat.Clone();
					}
					else
					{
						_paragraphFormat = new DocumentFormatInfo();
					}
					break;
				}
				lastElement = lastElement.Parent;
			}
		}
		#endregion

		#region AddContentElement
		/// <summary>
		/// Add content element
		/// </summary>
		/// <param name="newElement"></param>
		private void AddContentElement(DomElement newElement)
		{
			var elements = GetLastElements(true);
			DomElement lastElement = null;
			if (elements.Length > 0)
				lastElement = elements[elements.Length - 1];

			if (lastElement is DomDocument
				|| lastElement is DomHeader
				|| lastElement is DomFooter)
			{
				if (newElement is DomText
					|| newElement is DomImage
					|| newElement is DomObject
					|| newElement is DomShape
					|| newElement is DomShapeGroup)
				{
					var paragraph = new DomParagraph();
					if (lastElement.Elements.Count > 0)
						paragraph.IsTemplateGenerated = true;

					if (_paragraphFormat != null)
						paragraph.Format = _paragraphFormat;

					lastElement.AppendChild(paragraph);
					paragraph.Elements.Add(newElement);
					return;
				}
			}

			if (elements.Length == 0) return;
			var element = elements[elements.Length - 1];

			if (newElement != null && newElement.NativeLevel > 0)
			{
				for (var count = elements.Length - 1; count >= 0; count--)
				{
					if (elements[count].NativeLevel == newElement.NativeLevel)
					{
						for (var count2 = count; count2 < elements.Length; count2++)
						{
							var element2 = elements[count2];
							if (newElement is DomText
								|| newElement is DomImage
								|| newElement is DomObject
								|| newElement is DomShape
								|| newElement is DomShapeGroup
								|| newElement is DomField
								|| newElement is DomBookmark
								|| newElement is DomLineBreak)
							{
								if (newElement.NativeLevel == element2.NativeLevel)
								{
									if (element2 is DomTableRow
										|| element2 is DomTableCell
										|| element2 is DomField
										|| element2 is DomParagraph)
										continue;
								}
							}

							elements[count2].Locked = true;
						}
						break;
					}
				}
			}

			for (var count = elements.Length - 1; count >= 0; count--)
			{
				if (elements[count].Locked == false)
				{
					element = elements[count];
					if (element is DomImage)
						element.Locked = true;
					else
						break;
				}
			}
			if (element is DomTableRow)
			{
				// If the last element is table row 
				// can not contains any element , 
				// so need create a cell element.
				var tableCell = new DomTableCell { NativeLevel = element.NativeLevel };
				element.AppendChild(tableCell);
				if (newElement is DomTableRow)
				{
					tableCell.Elements.Add(newElement);
				}
				else
				{
					var cellParagraph = new DomParagraph
					{
						Format = _paragraphFormat.Clone(),
						NativeLevel = tableCell.NativeLevel
					};

					tableCell.AppendChild(cellParagraph);
					if (newElement != null)
						cellParagraph.AppendChild(newElement);
				}
			}
			else
			{
				if (newElement != null)
				{
					if (element is DomParagraph &&
						(newElement is DomParagraph
						 || newElement is DomTableRow))
					{
						// If both is paragraph , append new paragraph to the parent of old paragraph
						element.Locked = true;
						element.Parent.AppendChild(newElement);
					}
					else
					{
						element.AppendChild(newElement);
					}
				}
			}
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

		#region CombineTable
		/// <summary>
		/// Combine tables
		/// </summary>
		/// <param name="parentElement">The parent element</param>
		private void CombineTable(DomElement parentElement)
		{
			var result = new ArrayList();
			var rows = new ArrayList();
			var lastRowWidth = -1;
			DomTableRow lastRow = null;

			foreach (DomElement element in parentElement.Elements)
			{
				if (element is DomTableRow)
				{
					var tableRow = (DomTableRow)element;
					tableRow.Locked = false;
					var cellSettings = tableRow.CellSettings;

					if (cellSettings.Count == 0)
					{
						if (lastRow != null && lastRow.CellSettings.Count == tableRow.Elements.Count)
						{
							cellSettings = lastRow.CellSettings;
						}
					}

					if (cellSettings.Count == tableRow.Elements.Count)
					{
						for (var count = 0; count < tableRow.Elements.Count; count++)
							tableRow.Elements[count].Attributes = (AttributeList)cellSettings[count];
					}

					var isLastRow = tableRow.HasAttribute(Consts.Lastrow);

					if (isLastRow == false)
					{
						var index = parentElement.Elements.IndexOf(element);
						if (index == parentElement.Elements.Count - 1)
						{
							// If this element is the last element
							// then this row is the last row
							isLastRow = true;
						}
						else
						{
							var element2 = parentElement.Elements[index + 1];
							if (!(element2 is DomTableRow))
								// Next element is not row 
								isLastRow = true;
						}
					}

					// Split to table
					if (isLastRow)
					{
						// If current row mark the last row , then generate a new table
						rows.Add(tableRow);
						result.Add(CreateTable(rows));
						lastRowWidth = -1;
					}
					else
					{
						var width = 0;
						if (tableRow.HasAttribute(Consts.TrwWidth))
						{
							width = tableRow.Attributes[Consts.TrwWidth];
							if (tableRow.HasAttribute(Consts.TrwWidthA))
								width = width - tableRow.Attributes[Consts.TrwWidthA];
						}
						else
						{
							foreach (DomTableCell tableCell in tableRow.Elements)
							{
								if (tableCell.HasAttribute(Consts.Cellx))
									width = Math.Max(width, tableCell.Attributes[Consts.Cellx]);
							}
						}
						if (lastRowWidth > 0 && lastRowWidth != width)
						{
							// If row's width is change , then can consider multi-table combin
							// then split and generate new table
							if (rows.Count > 0)
								result.Add(CreateTable(rows));
						}
						lastRowWidth = width;
						rows.Add(tableRow);
					}
					lastRow = tableRow;
				}
				else if (element is DomTableCell)
				{
					lastRow = null;
					CombineTable(element);
					if (rows.Count > 0)
					{
						result.Add(CreateTable(rows));
					}
					result.Add(element);
					lastRowWidth = -1;
				}
				else
				{
					lastRow = null;
					CombineTable(element);
					if (rows.Count > 0)
						result.Add(CreateTable(rows));
					result.Add(element);
					lastRowWidth = -1;
				}
			}

			if (rows.Count > 0)
				result.Add(CreateTable(rows));

			parentElement.Locked = false;
			parentElement.Elements.Clear();

			foreach (DomElement element in result)
				parentElement.AppendChild(element);
		}
		#endregion

		#region CreateTable
		/// <summary>
		/// Create table
		/// </summary>
		/// <param name="rows">table rows</param>
		/// <returns>new table</returns>
		private DomTable CreateTable(ArrayList rows)
		{
			if (rows.Count > 0)
			{
				var table = new DomTable();
				var index = 0;
				foreach (DomTableRow row in rows)
				{
					row.RowIndex = index;
					index++;
					table.AppendChild(row);
				}
				rows.Clear();
				foreach (DomTableRow tableRow in table.Elements)
				{
					foreach (DomTableCell cell in tableRow.Elements)
						CombineTable(cell);
				}
				return table;
			}
			throw new ArgumentException("rows");
		}
		#endregion

		#region UpdateTableCells
		private void UpdateTableCells(DomTable table, bool fixTableCellSize)
		{
			// Number of table column
			var columns = 0;
			// Flag of cell merge
			var merge = false;
			// Right position of all cells
			var rights = new ArrayList();

			// Right position of table
			var tableLeft = 0;
			for (var count = table.Elements.Count - 1; count >= 0; count--)
			{
				var tableRow = (DomTableRow)table.Elements[count];
				if (tableRow.Elements.Count == 0)
					table.Elements.RemoveAt(count);
			}

			if (table.Elements.Count == 0)
				Debug.WriteLine("");

			foreach (DomTableRow tableRow in table.Elements)
			{
				var lastCellX = 0;

				columns = Math.Max(columns, tableRow.Elements.Count);
				if (tableRow.HasAttribute(Consts.Irow))
				{
					tableRow.RowIndex = tableRow.Attributes[Consts.Irow];
				}
				tableRow.IsLastRow = tableRow.HasAttribute(Consts.Lastrow);
				tableRow.Header = tableRow.HasAttribute(Consts.Trhdr);
				// Read row height
				if (tableRow.HasAttribute(Consts.Trrh))
				{
					tableRow.Height = tableRow.Attributes[Consts.Trrh];
					if (tableRow.Height == 0)
						tableRow.Height = DefaultRowHeight;
					else if (tableRow.Height < 0)
						tableRow.Height = -tableRow.Height;
				}
				else
					tableRow.Height = DefaultRowHeight;

				// Read default padding of cell
				tableRow.PaddingLeft = tableRow.HasAttribute(Consts.Trpaddl) ? tableRow.Attributes[Consts.Trpaddl] : int.MinValue;
				tableRow.PaddingTop = tableRow.HasAttribute(Consts.Trpaddt) ? tableRow.Attributes[Consts.Trpaddt] : int.MinValue;
				tableRow.PaddingRight = tableRow.HasAttribute(Consts.Trpaddr) ? tableRow.Attributes[Consts.Trpaddr] : int.MinValue;
				tableRow.PaddingBottom = tableRow.HasAttribute(Consts.Trpaddb) ? tableRow.Attributes[Consts.Trpaddb] : int.MinValue;

				if (tableRow.HasAttribute(Consts.Trleft))
					tableLeft = tableRow.Attributes[Consts.Trleft];

				if (tableRow.HasAttribute(Consts.Trcbpat))
					tableRow.Format.BackColor = ColorTable.GetColor(
						tableRow.Attributes[Consts.Trcbpat],
						Color.Transparent);

				var widthCount = 0;
				foreach (DomTableCell cell in tableRow.Elements)
				{
					// Set cell's dispaly format
					if (cell.HasAttribute(Consts.Clvmgf))
						merge = true;

					if (cell.HasAttribute(Consts.Clvmrg))
						merge = true;

					cell.PaddingLeft = cell.HasAttribute(Consts.Clpadl) ? cell.Attributes[Consts.Clpadl] : int.MinValue;
					cell.PaddingRight = cell.HasAttribute(Consts.Clpadr) ? cell.Attributes[Consts.Clpadr] : int.MinValue;
					cell.PaddingTop = cell.HasAttribute(Consts.Clpadt) ? cell.Attributes[Consts.Clpadt] : int.MinValue;
					cell.PaddingBottom = cell.HasAttribute(Consts.Clpadb) ? cell.Attributes[Consts.Clpadb] : int.MinValue;

					// Whether display border line
					cell.Format.LeftBorder = cell.HasAttribute(Consts.Clbrdrl);
					cell.Format.TopBorder = cell.HasAttribute(Consts.Clbrdrt);
					cell.Format.RightBorder = cell.HasAttribute(Consts.Clbrdrr);
					cell.Format.BottomBorder = cell.HasAttribute(Consts.Clbrdrb);

					if (cell.HasAttribute(Consts.Brdrcf))
					{
						cell.Format.BorderColor = ColorTable.GetColor(
							cell.GetAttributeValue(Consts.Brdrcf, 1),
							Color.Black);
					}

					for (var count = cell.Attributes.Count - 1; count >= 0; count--)
					{
						var name3 = cell.Attributes.GetItem(count).Name;
						if (name3 == Consts.Brdrtbl
							|| name3 == Consts.Brdrnone
							|| name3 == Consts.Brdrnil)
						{
							for (var count2 = count - 1; count2 >= 0; count2--)
							{
								var name2 = cell.Attributes.GetItem(count2).Name;
								if (name2 == Consts.Clbrdrl)
								{
									cell.Format.LeftBorder = false;
									break;
								}

								if (name2 == Consts.Clbrdrt)
								{
									cell.Format.TopBorder = false;
									break;
								}

								if (name2 == Consts.Clbrdrr)
								{
									cell.Format.RightBorder = false;
									break;
								}

								if (name2 == Consts.Clbrdrb)
								{
									cell.Format.BottomBorder = false;
									break;
								}
							}
						}
					}

					// Vertial alignment
					if (cell.HasAttribute(Consts.Clvertalt))
						cell.VerticalAlignment = RtfVerticalAlignment.Top;
					else if (cell.HasAttribute(Consts.Clvertalc))
						cell.VerticalAlignment = RtfVerticalAlignment.Middle;
					else if (cell.HasAttribute(Consts.Clvertalb))
						cell.VerticalAlignment = RtfVerticalAlignment.Bottom;

					// Background color
					cell.Format.BackColor = cell.HasAttribute(Consts.Clcbpat) ? ColorTable.GetColor(cell.Attributes[Consts.Clcbpat], Color.Transparent) : Color.Transparent;
					if (cell.HasAttribute(Consts.Clcfpat))
						cell.Format.BorderColor = ColorTable.GetColor(cell.Attributes[Consts.Clcfpat], Color.Black);

					// Cell's width
					var cellWidth = 2763; // cell's default with is 2763 Twips(570 Document)
					if (cell.HasAttribute(Consts.Cellx))
					{
						cellWidth = cell.Attributes[Consts.Cellx] - lastCellX;
						if (cellWidth < 100)
							cellWidth = 100;
					}

					var right = lastCellX + cellWidth;
					// fix cell's right position , if this position is very near with another cell's 
					// right position( less then 45 twips or 3 pixel), then consider these two position
					// is the same , this can decrease number of table columns
					foreach (var t in rights)
					{
						if (Math.Abs(right - (int)t) < 45)
						{
							right = (int)t;
							cellWidth = right - lastCellX;
							break;
						}
					}

					cell.Left = lastCellX;
					cell.Width = cellWidth;

					widthCount += cellWidth;

					if (rights.Contains(right) == false)
					{
						// becase of convert twips to unit of document may cause truncation error.
						// This may cause rights.Contains mistake . so scale cell's with with 
						// native twips unit , after all computing , convert to unit of document.
						rights.Add(right);
					}
					lastCellX = lastCellX + cellWidth;
				} //foreach
				tableRow.Width = widthCount;
			} //foreach

			if (rights.Count == 0)
			{
				// can not detect cell's width , so consider set cell's width
				// automatic, then set cell's default width.
				var cols = 1;
				foreach (DomTableRow row in table.Elements)
					cols = Math.Max(cols, row.Elements.Count);

				var w = ClientWidth / cols;

				for (var count = 0; count < cols; count++)
					rights.Add(count * w + w);
			}

			// Computing cell's rowspan and colspan , number of rights array is the number of table columns.

			rights.Add(0);
			rights.Sort();

			// Add table column instance
			for (var count = 1; count < rights.Count; count++)
			{
				var col = new DomTableColumn { Width = (int)rights[count] - (int)rights[count - 1] };
				table.Columns.Add(col);
			}

			for (var rowIndex = 1; rowIndex < table.Elements.Count; rowIndex++)
			{
				var row = (DomTableRow)table.Elements[rowIndex];
				for (var colIndex = 0; colIndex < row.Elements.Count; colIndex++)
				{
					var cell = (DomTableCell)row.Elements[colIndex];
					if (cell.Width == 0)
					{
						// If current cell not special width , then use the width of cell which 
						// in the same colum and in the last row
						var preRow = (DomTableRow)table.Elements[rowIndex - 1];
						if (preRow.Elements.Count > colIndex)
						{
							var preCell = (DomTableCell)preRow.Elements[colIndex];
							cell.Left = preCell.Left;
							cell.Width = preCell.Width;
							CopyStyleAttribute(cell, preCell.Attributes);
						}
					}
				}
			}

			if (merge == false)
			{
				// If not detect cell merge , maby exist cell merge in the same row
				foreach (DomTableRow row in table.Elements)
				{
					if (row.Elements.Count < table.Columns.Count)
					{
						// If number of row's cells not equals the number of table's columns
						// then exist cell merge.
						merge = true;
						break;
					}
				}
			}

			if (merge)
			{
				// Detect cell merge , begin merge operation

				// Because of in rtf format,cell which merged by another cell in the same row , 
				// does no written in rtf text , so delay create those cell instance .
				foreach (DomTableRow row in table.Elements)
				{
					if (row.Elements.Count != table.Columns.Count)
					{
						// If number of row's cells not equals number of table's columns ,
						// then consider there are hanppend  horizontal merge.
						var cells = row.Elements.ToArray();
						foreach (var domElement in cells)
						{
							var cell = (DomTableCell)domElement;
							var index = rights.IndexOf(cell.Left);
							var index2 = rights.IndexOf(cell.Left + cell.Width);
							var intColSpan = index2 - index;
							// detect vertical merge
							var verticalMerge = cell.HasAttribute(Consts.Clvmrg);

							if (verticalMerge == false)
							{
								// If this cell does not merged by another cell abover , 
								// then set colspan
								cell.ColSpan = intColSpan;
							}

							if (row.Elements.LastElement == cell)
							{
								cell.ColSpan = table.Columns.Count - row.Elements.Count + 1;
								intColSpan = cell.ColSpan;
							}

							for (var count = 0; count < intColSpan - 1; count++)
							{
								var newCell = new DomTableCell { Attributes = cell.Attributes.Clone() };
								row.Elements.Insert(row.Elements.IndexOf(cell) + 1, newCell);
								if (verticalMerge)
								{
									// This cell has been merged.
									newCell.Attributes[Consts.Clvmrg] = 1;
									newCell.OverrideCell = cell;
								}
							}
						}

						if (row.Elements.Count != table.Columns.Count)
						{
							// If the last cell has been merged. then supply new cells.
							var lastCell = (DomTableCell)row.Elements.LastElement;
							if (lastCell == null)
								Console.WriteLine("");


							for (var count = row.Elements.Count; count < rights.Count; count++)
							{
								var newCell = new DomTableCell();
								if (lastCell != null) CopyStyleAttribute(newCell, lastCell.Attributes);
								row.Elements.Add(newCell);
							}
						}
					}
				}

				// Set cell's vertial merge.
				foreach (DomTableRow tableRow in table.Elements)
				{
					foreach (DomTableCell tableCell in tableRow.Elements)
					{
						if (tableCell.HasAttribute(Consts.Clvmgf) == false)
						{
							//if this cell does not mark vertial merge , then next cell
							continue;
						}
						// if this cell mark vertial merge.
						var colIndex = tableRow.Elements.IndexOf(tableCell);
						for (var rowIndex = table.Elements.IndexOf(tableRow) + 1;
							rowIndex < table.Elements.Count;
							rowIndex++)
						{
							var row2 = (DomTableRow)table.Elements[rowIndex];
							if (colIndex >= row2.Elements.Count)
							{
								Console.Write("");
							}
							var cell2 = (DomTableCell)row2.Elements[colIndex];
							if (cell2.HasAttribute(Consts.Clvmrg))
							{
								if (cell2.OverrideCell != null)
									// If this cell has been merge by another cell( must in the same row )
									// then break the circle
									break;

								// Increase vertial merge.
								tableCell.RowSpan++;
								cell2.OverrideCell = tableCell;
							}
							else
								// if this cell not mark merged by another cell , then break the circel
								break;
						}
					}
				}

				// Set cell's OverridedCell information
				foreach (DomTableRow tableRow in table.Elements)
				{
					foreach (DomTableCell tableCell in tableRow.Elements)
					{
						if (tableCell.RowSpan > 1 || tableCell.ColSpan > 1)
						{
							for (var rowIndex = 1; rowIndex <= tableCell.RowSpan; rowIndex++)
							{
								for (var colIndex = 1; colIndex <= tableCell.ColSpan; colIndex++)
								{
									var r = table.Elements.IndexOf(tableRow) + rowIndex - 1;
									var c = tableRow.Elements.IndexOf(tableCell) + colIndex - 1;
									var cell2 = (DomTableCell)table.Elements[r].Elements[c];
									if (tableCell != cell2)
									{
										cell2.OverrideCell = tableCell;
									}
								}
							}
						}
					}
				}
			}

			if (fixTableCellSize)
			{
				// Fix table's left position use the first table column
				if (table.Columns.Count > 0)
					((DomTableColumn)table.Columns[0]).Width -= tableLeft;
			}
		}
		#endregion

		#region CopyStyleAttribute
		private void CopyStyleAttribute(DomTableCell cell, AttributeList table)
		{
			var attrs = table.Clone();
			attrs.Remove(Consts.Clvmgf);
			attrs.Remove(Consts.Clvmrg);
			cell.Attributes = attrs;
		}
		#endregion

		#region ToString
		public override string ToString()
		{
			return "RTFDocument:" + Info.Title;
		}
		#endregion

		#region ApplyText
		private bool ApplyText(TextContainer textContainer, Reader reader, DocumentFormatInfo format)
		{
			if (textContainer.HasContent)
			{
				var text = textContainer.Text;
				textContainer.Clear();

				var image = (DomImage)GetLastElement(typeof(DomImage));
				if (image != null && image.Locked == false)
				{
					image.Data = HexToBytes(text);
					image.Format = format.Clone();
					image.Width = image.DesiredWidth * image.ScaleX / 100;
					image.Height = image.DesiredHeight * image.ScaleY / 100;
					image.Locked = true;
					if (reader.TokenType != RtfTokenType.GroupEnd)
					{
						ReadToEndGround(reader);
					}
					return true;
				}
				if (format.ReadText && _startContent)
				{
					var txt = new DomText { NativeLevel = textContainer.Level, Format = format.Clone() };
					if (txt.Format.Align == RtfAlignment.Justify)
						txt.Format.Align = RtfAlignment.Left;
					txt.Text = text;
					AddContentElement(txt);
				}
			}
			return false;
		}
		#endregion

		#region ReadToEndGround
		/// <summary>
		/// Read data , until at the front of the end token belong the current level.
		/// </summary>
		/// <param name="reader"></param>
		private void ReadToEndGround(Reader reader)
		{
			reader.ReadToEndGround();
		}
		#endregion

		#region ReadListOverrideTable
		private void ReadListOverrideTable(Reader reader)
		{
			ListOverrideTable = new ListOverrideTable();
			while (reader.ReadToken() != null)
			{
				if (reader.TokenType == RtfTokenType.GroupEnd)
					break;

				if (reader.TokenType == RtfTokenType.GroupStart)
				{
					ListOverride record = null;
					while (reader.ReadToken() != null)
					{
						if (reader.TokenType == RtfTokenType.GroupEnd)
						{
							break;
						}
						if (reader.CurrentToken.Key == Consts.ListOverride)
						{
							record = new ListOverride();
							ListOverrideTable.Add(record);
							continue;
						}

						if (record == null)
							continue;

						switch (reader.CurrentToken.Key)
						{
							case Consts.ListId:
								record.ListId = reader.CurrentToken.Param;
								break;

							case Consts.ListOverrideCount:
								record.ListOverrideCount = reader.CurrentToken.Param;
								break;

							case Consts.Ls:
								record.Id = reader.CurrentToken.Param;
								break;
						}
					}
				}
			}
		}
		#endregion

		#region ReadListTable
		private void ReadListTable(Reader reader)
		{
			ListTable = new ListTable();
			while (reader.ReadToken() != null)
			{
				if (reader.TokenType == RtfTokenType.GroupEnd)
					break;

				if (reader.TokenType == RtfTokenType.GroupStart)
				{
					var firstRead = true;
					RtfList currentList = null;
					var level = reader.Level;
					while (reader.ReadToken() != null)
					{
						if (reader.TokenType == RtfTokenType.GroupEnd)
						{
							if (reader.Level < level)
							{
								break;
							}
						}
						else if (reader.TokenType == RtfTokenType.GroupStart)
						{
							// if meet nested level , then ignore
							//reader.ReadToken();
							//ReadToEndGround(reader);
							//reader.ReadToken();
						}
						if (firstRead)
						{
							if (reader.CurrentToken.Key != "list")
							{
								// 不是以list开头，忽略掉
								ReadToEndGround(reader);
								reader.ReadToken();
								break;
							}
							currentList = new RtfList();
							ListTable.Add(currentList);
							firstRead = false;
						}

						switch (reader.CurrentToken.Key)
						{
							case "listtemplateid":
								currentList.ListTemplateId = reader.CurrentToken.Param;
								break;

							case "listid":
								currentList.ListId = reader.CurrentToken.Param;
								break;

							case "listhybrid":
								currentList.ListHybrid = true;
								break;

							case "levelfollow":
								currentList.LevelFollow = reader.CurrentToken.Param;
								break;

							case "levelstartat":
								currentList.LevelStartAt = reader.CurrentToken.Param;
								break;

							case "levelnfc":
								if (currentList.LevelNfc == RtfLevelNumberType.None)
									currentList.LevelNfc = (RtfLevelNumberType)reader.CurrentToken.Param;
								break;

							case "levelnfcn":
								if (currentList.LevelNfc == RtfLevelNumberType.None)
									currentList.LevelNfc = (RtfLevelNumberType)reader.CurrentToken.Param;
								break;

							case "leveljc":
								currentList.LevelJc = reader.CurrentToken.Param;
								break;

							case "leveltext":
								if (string.IsNullOrEmpty(currentList.LevelText))
								{
									var text = ReadInnerText(reader, true);
									if (text != null && text.Length > 2)
									{
										int len = text[0];
										len = Math.Min(len, text.Length - 1);
										text = text.Substring(1, len);
									}
									currentList.LevelText = text;
								}
								break;

							case "f":
								currentList.FontName = FontTable.GetFontName(reader.CurrentToken.Param);
								break;
						}
					}
				}
			}
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
							ReadToEndGround(reader);
							reader.ReadToken();
						}
						else if (reader.Keyword == "f" && reader.HasParam)
							index = reader.Parameter;
						else if (reader.Keyword == "fnil")
						{
							name = Control.DefaultFont.Name;
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
							name = Control.DefaultFont.Name;

						var font = new Font(index, name) { Charset = charset, NilFlag = nilFlag };
						FontTable.Add(font);
					}
				}
			}
		}
		#endregion

		#region ReadColorTable
		/// <summary>
		/// Read color table
		/// </summary>
		/// <param name="reader"></param>
		private void ReadColorTable(Reader reader)
		{
			ColorTable.Clear();
			ColorTable.CheckValueExistWhenAdd = false;
			var r = -1;
			var g = -1;
			var b = -1;
			while (reader.ReadToken() != null)
			{
				if (reader.TokenType == RtfTokenType.GroupEnd)
				{
					break;
				}
				switch (reader.Keyword)
				{
					case "red":
						r = reader.Parameter;
						break;
					case "green":
						g = reader.Parameter;
						break;
					case "blue":
						b = reader.Parameter;
						break;
					case ";":
						if (r >= 0 && g >= 0 && b >= 0)
						{
							var c = Color.FromArgb(255, r, g, b);
							ColorTable.Add(c);
							r = -1;
							g = -1;
							b = -1;
						}
						break;
				}
			}
			if (r >= 0 && g >= 0 && b >= 0)
			{
				// Read the last color
				var c = Color.FromArgb(255, r, g, b);
				ColorTable.Add(c);
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

		#region ReadDomObject
		/// <summary>
		/// Read a rtf emb object
		/// </summary>
		/// <param name="reader">reader</param>
		/// <param name="format">format</param>
		/// <returns>rtf emb object instance</returns>
		private void ReadDomObject(Reader reader, DocumentFormatInfo format)
		{
			var domObject = new DomObject { NativeLevel = reader.Level };
			AddContentElement(domObject);
			var levelBack = reader.Level;

			while (reader.ReadToken() != null)
			{
				if (reader.Level < levelBack)
					break;
				if (reader.TokenType == RtfTokenType.GroupStart)
					continue;
				if (reader.TokenType == RtfTokenType.GroupEnd)
					continue;
				if (reader.Level == domObject.NativeLevel + 1
					&& reader.Keyword.StartsWith("attribute_"))
					domObject.CustomAttributes[reader.Keyword] = ReadInnerText(reader, true);

				switch (reader.Keyword)
				{
					case Consts.Objautlink:
						domObject.Type = RtfObjectType.AutLink;
						break;

					case Consts.Objclass:
						domObject.ClassName = ReadInnerText(reader, true);
						break;

					case Consts.Objdata:
						var data = ReadInnerText(reader, true);
						domObject.Content = HexToBytes(data);
						break;

					case Consts.Objemb:
						domObject.Type = RtfObjectType.Emb;
						break;

					case Consts.Objh:
						domObject.Height = reader.Parameter;
						break;

					case Consts.Objhtml:
						domObject.Type = RtfObjectType.Html;
						break;

					case Consts.Objicemb:
						domObject.Type = RtfObjectType.Icemb;
						break;

					case Consts.Objlink:
						domObject.Type = RtfObjectType.Link;
						break;

					case Consts.Objname:
						domObject.Name = ReadInnerText(reader, true);
						break;

					case Consts.Objocx:
						domObject.Type = RtfObjectType.Ocx;
						break;

					case Consts.Objpub:
						domObject.Type = RtfObjectType.Pub;
						break;

					case Consts.Objsub:
						domObject.Type = RtfObjectType.Sub;
						break;

					case Consts.Objtime:
						break;

					case Consts.Objw:
						domObject.Width = reader.Parameter;
						break;

					case Consts.Objscalex:
						domObject.ScaleX = reader.Parameter;
						break;

					case Consts.Objscaley:
						domObject.ScaleY = reader.Parameter;
						break;

					case Consts.Result:
						var result = new ElementContainer { Name = Consts.Result };
						domObject.AppendChild(result);
						Load(reader, format);
						result.Locked = true;
						break;
				}
			}
			domObject.Locked = true;
		}
		#endregion

		#region ReadDomField
		/// <summary>
		/// Read field
		/// </summary>
		/// <param name="reader"></param>
		/// <param name="format"></param>
		/// <returns></returns>
		private void ReadDomField(
			Reader reader,
			DocumentFormatInfo format)
		{
			var field = new DomField { NativeLevel = reader.Level };
			AddContentElement(field);
			var levelBack = reader.Level;
			while (reader.ReadToken() != null)
			{
				if (reader.Level < levelBack)
					break;

				if (reader.TokenType == RtfTokenType.GroupStart)
				{
				}
				else if (reader.TokenType == RtfTokenType.GroupEnd)
				{
				}
				else
				{
					switch (reader.Keyword)
					{
						case Consts.Flddirty:
							field.Method = RtfDomFieldMethod.Dirty;
							break;

						case Consts.Fldedit:
							field.Method = RtfDomFieldMethod.Edit;
							break;

						case Consts.Fldlock:
							field.Method = RtfDomFieldMethod.Lock;
							break;

						case Consts.Fldpriv:
							field.Method = RtfDomFieldMethod.Priv;
							break;

						case Consts.Fldrslt:
							var result = new ElementContainer { Name = Consts.Fldrslt };
							field.AppendChild(result);
							Load(reader, format);
							result.Locked = true;
							break;

						case Consts.Fldinst:
							var inst = new ElementContainer { Name = Consts.Fldinst };
							field.AppendChild(inst);
							Load(reader, format);
							inst.Locked = true;
							var txt = inst.InnerText;
							if (txt != null)
							{
								var index = txt.IndexOf(Consts.Hyperlink, StringComparison.Ordinal);
								if (index >= 0)
								{
									var index1 = txt.IndexOf('\"', index);
									if (index1 > 0 && txt.Length > index1 + 2)
									{
										var index2 = txt.IndexOf('\"', index1 + 2);
										if (index2 > index1)
										{
											var link = txt.Substring(index1 + 1, index2 - index1 - 1);
											if (format.Parent != null)
											{
												if (link.StartsWith("_Toc"))
													link = "#" + link;
												format.Parent.Link = link;
											}
										}
									}
								}
							}

							break;
					}
				}
			}
			field.Locked = true;
			//return field;
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

		#region ToDomString
		public override string ToDomString()
		{
			var builder = new StringBuilder();
			builder.Append(ToString());
			builder.Append(Environment.NewLine + "   Info");
			foreach (var item in Info.StringItems)
				builder.Append(Environment.NewLine + "      " + item);

			builder.Append(Environment.NewLine + "   ColorTable(" + ColorTable.Count + ")");

			for (var count = 0; count < ColorTable.Count; count++)
			{
				var c = ColorTable[count];
				builder.Append(Environment.NewLine + "      " + count + ":" + c.R + " " + c.G + " " + c.B);
			}

			builder.Append(Environment.NewLine + "   FontTable(" + FontTable.Count + ")");

			foreach (Font font in FontTable)
				builder.Append(Environment.NewLine + "      " + font);

			if (ListTable.Count > 0)
			{
				builder.Append(Environment.NewLine + "   ListTable(" + ListTable.Count + ")");
				foreach (var list in ListTable)
					builder.Append(Environment.NewLine + "      " + list);
			}

			if (ListOverrideTable.Count > 0)
			{
				builder.Append(Environment.NewLine + "   ListOverrideTable(" + ListOverrideTable.Count + ")");
				foreach (var list in ListOverrideTable)
					builder.Append(Environment.NewLine + "      " + list);
			}
			builder.Append(Environment.NewLine + "   -----------------------");

			if (string.IsNullOrEmpty(HtmlContent) == false)
			{
				builder.Append(Environment.NewLine + "   HTMLContent:" + HtmlContent);
				builder.Append(Environment.NewLine + "   -----------------------");
			}

			ToDomString(Elements, builder, 1);
			return builder.ToString();
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
            //int? fontIndex = null;
            //float? fontSize = null;
            //var spanTagWritten = false;
            var encoding = _defaultEncoding;

            while (reader.ReadToken() != null)
			{
				switch (reader.Keyword)
				{
					//case Consts.Fonttbl:
					//	// Read font table
					//	ReadFontTable(reader);
					//	break;

                    //case Consts.F:
                    //    if (reader.HasParam)
                    //    {
                    //        if (spanTagWritten && fontIndex.HasValue)
                    //        {
                    //            stringBuilder.Append("</span>");
                    //            spanTagWritten = false;
                    //            fontSize = null;
                    //            encoding = _defaultEncoding;
                    //        }

                    //        fontIndex = reader.Parameter;
                    //    }
                    //    break;

                    //case Consts.Fs:
                    //    if (reader.HasParam)
                    //        fontSize = reader.Parameter / 2.0f;
                    //    break;

                    case Consts.HtmlRtf:
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
				            }
				            htmlState = true;
				        }
				        break;

					case Consts.HtmlTag:
						if (reader.InnerReader.Peek() == ' ')
							reader.InnerReader.Read();

						var text = ReadInnerText(reader, null, true, false, true);

                        //if (spanTagWritten && text.StartsWith("<span"))
                        //{
                        //    stringBuilder.Append("</span>");
                        //    spanTagWritten = false;
                        //    fontIndex = null;
                        //    fontSize = null;
                        //    encoding = _defaultEncoding;
                        //}

                        if (!string.IsNullOrEmpty(text))
                            stringBuilder.Append(text);
                        break;

                    default:

						switch (reader.TokenType)
						{
							case RtfTokenType.Control:
								if (!htmlState)
								{
                                    //if (spanTagWritten && reader.Keyword != "'")
                                    //{
                                    //    stringBuilder.Append("</span>");
                                    //    spanTagWritten = false;
                                    //    fontIndex = null;
                                    //    fontSize = null;
                                    //    encoding = _defaultEncoding;
                                    //}

                                    switch (reader.Keyword)
									{
										case "'":

                                            //if (FontTable != null && fontIndex.HasValue && fontIndex <= FontTable.Count)
                                            //{
                                            //    // <span style = 'font-size:12.0pt;font-family:"Arial",sans-serif' >
                                            //    var font = FontTable[fontIndex.Value];
                                            //    if (!spanTagWritten)
                                            //    {
                                            //        stringBuilder.Append("<span style = 'font-family:\"" + font.Name + "\";");
                                            //        if (fontSize.HasValue)
                                            //            stringBuilder.Append("font-size:" + fontSize + "pt");
                                            //        stringBuilder.Append("'>");
                                            //        spanTagWritten = true;
                                            //        encoding = font.Encoding ?? _defaultEncoding;
                                            //    }
                                            //}

									        // Convert HEX value directly when we have a single byte charset
									        if (_defaultEncoding.IsSingleByte )
									        {
									            if (string.IsNullOrEmpty(hexBuffer))
									                hexBuffer = reader.CurrentToken.Hex;

                                                var buff = new[] { byte.Parse(hexBuffer, NumberStyles.HexNumber) };
                                                hexBuffer = string.Empty;
                                                stringBuilder.Append(encoding.GetString(buff));
                                            }
									        else
									        {
									            // If we have a double byte charset like chinese then store the value and wait for the next HEX value
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

                                            //case "u":
                                            //    stringBuilder.Append(HttpUtility.UrlDecode("*", _defaultEncoding));
                                            //    break;
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

										case Consts.U:
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