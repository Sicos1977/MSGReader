using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF document writer
    /// </summary>
    internal sealed class DocumentWriter
    {
        #region Fields
        private bool _firstParagraph = true;
        private DocumentFormatInfo _lastParagraphInfo;
        #endregion

        #region Properties
        // ReSharper disable MemberCanBePrivate.Global
        /// <summary>
        /// Base writer
        /// </summary>
        public Writer Writer { get; set; }

        /// <summary>
        /// Information about this Rtf document
        /// </summary>
        public Hashtable Info { get; private set; }

        /// <summary>
        /// Rtf font table
        /// </summary>
        public Table FontTable { get; private set; }

        public ListTable ListTable { get; set; }

        public ListOverrideTable ListOverrideTable { get; set; }

        /// <summary>
        /// Rtf color table
        /// </summary>
        public ColorTable ColorTable { get; private set; }

        /// <summary>
        /// System collectiong document's information , maby generating
        /// font table and color table , not writting content.
        /// </summary>
        public bool CollectionInfo { get; set; }

        /// <summary>
        /// How many nested groups do we have
        /// </summary>
        public int GroupLevel
        {
            get { return Writer.GroupLevel; }
        }

        /// <summary>
        /// When debug mode is turned on, raw information about the Rtf file is written
        /// </summary>
        public bool DebugMode { get; set; }
        // ReSharper restore MemberCanBePrivate.Global
        #endregion

        #region Constructors
        /// <summary>
        /// Initialize instance
        /// </summary>
        public DocumentWriter()
        {
            Info = new Hashtable();
            DebugMode = true;
            CollectionInfo = true;
            ColorTable = new ColorTable();
            ListOverrideTable = new ListOverrideTable();
            ListTable = new ListTable();
            FontTable = new Table();
            ColorTable.CheckValueExistWhenAdd = true;
        }

        /// <summary>
        /// Initialize instance with a text writer
        /// </summary>
        /// <param name="writer"></param>
        public DocumentWriter(TextWriter writer)
        {
            Info = new Hashtable();
            DebugMode = true;
            CollectionInfo = true;
            ColorTable = new ColorTable();
            ListOverrideTable = new ListOverrideTable();
            ListTable = new ListTable();
            FontTable = new Table();
            ColorTable.CheckValueExistWhenAdd = true;
            Open(writer);
        }

        /// <summary>
        /// Initialize instance from a file
        /// </summary>
        /// <param name="fileName"></param>
        public DocumentWriter(string fileName)
        {
            Info = new Hashtable();
            DebugMode = true;
            CollectionInfo = true;
            ColorTable = new ColorTable();
            ListOverrideTable = new ListOverrideTable();
            ListTable = new ListTable();
            FontTable = new Table();
            ColorTable.CheckValueExistWhenAdd = true;
            // ReSharper disable once DoNotCallOverridableMethodsInConstructor
            Open(fileName);
        }

        /// <summary>
        /// Initialize instance from a stream
        /// </summary>
        /// <param name="stream"></param>
        public DocumentWriter(Stream stream)
        {
            Info = new Hashtable();
            DebugMode = true;
            CollectionInfo = true;
            ColorTable = new ColorTable();
            ListOverrideTable = new ListOverrideTable();
            ListTable = new ListTable();
            FontTable = new Table();
            ColorTable.CheckValueExistWhenAdd = true;
            var writer = new StreamWriter(stream, Encoding.ASCII);
            // ReSharper disable once DoNotCallOverridableMethodsInConstructor
            Open(writer);
        }
        #endregion

        #region Open
        // ReSharper disable MemberCanBeProtected.Global
        /// <summary>
        /// Open Rtf file from a textwriter
        /// </summary>
        /// <param name="writer"></param>
        public void Open(TextWriter writer)
        {
            Writer = new Writer(writer) {Indent = false};
        }

        /// <summary>
        /// Open Rtf file from a file
        /// </summary>
        /// <param name="fileName"></param>
        public void Open(string fileName)
        {
            Writer = new Writer(fileName) {Indent = false};
        }
        // ReSharper restore MemberCanBeProtected.Global
        #endregion

        #region Close
        public void Close()
        {
            Writer.Close();
        }
        #endregion

        #region WriteGroup
        public void WriteStartGroup()
        {
            if (CollectionInfo == false)
                Writer.WriteStartGroup();
        }

        public void WriteEndGroup()
        {
            if (CollectionInfo == false)
                Writer.WriteEndGroup();
        }
        #endregion

        #region WriteKeyword
        /// <summary>
        /// Write rtf keyword
        /// </summary>
        /// <param name="keyword">keyword</param>
        // ReSharper disable once MemberCanBePrivate.Global
        public void WriteKeyword(string keyword)
        {
            if (CollectionInfo == false)
                Writer.WriteKeyword(keyword);
        }
        
        /// <summary>
        /// Write rtf keyword
        /// </summary>
        /// <param name="keyWord"></param>
        /// <param name="ext">Keyword is external</param>
        public void WriteKeyword(string keyWord, bool ext)
        {
            if (CollectionInfo == false)
                Writer.WriteKeyword(keyWord, ext);
        }
        #endregion

        #region WriteRaw
        /// <summary>
        /// Write raw text
        /// </summary>
        /// <param name="text"></param>
        public void WriteRaw(string text)
        {
            if (CollectionInfo == false)
            {
                if (text != null)
                    Writer.WriteRaw(text);
            }
        }
        #endregion

        #region WriteBorderLineDashStyle
        /// <summary>
        /// Write the style that is used by the borderline
        /// </summary>
        /// <param name="style"></param>
        public void WriteBorderLineDashStyle(DashStyle style)
        {
            if (CollectionInfo == false)
            {
                switch (style)
                {
                    case DashStyle.Dot:
                        WriteKeyword("brdrdot");
                        break;

                    case DashStyle.DashDot:
                        WriteKeyword("brdrdashd");
                        break;

                    case DashStyle.DashDotDot:
                        WriteKeyword("brdrdashdd");
                        break;

                    case DashStyle.Dash:
                        WriteKeyword("brdrdash");
                        break;

                    default:
                        WriteKeyword("brdrs");
                        break;
                }
            }
        }
        #endregion

        #region WriteStartEndDocument
        /// <summary>
        /// Write the start of the document
        /// </summary>
        public void WriteStartDocument()
        {
            _lastParagraphInfo = null;
            _firstParagraph = true;
            if (CollectionInfo)
            {
                Info.Clear();
                FontTable.Clear();
                ColorTable.Clear();
                FontTable.Add(Control.DefaultFont.Name);
            }
            else
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword(Consts.Rtf);
                Writer.WriteKeyword("ansi");
                Writer.WriteKeyword("ansicpg" + Writer.Encoding.CodePage);
                
                // Write document information
                if (Info.Count > 0)
                {
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("info");
                    foreach (string key in Info.Keys)
                    {
                        Writer.WriteStartGroup();

                        var value = Info[key];
                        if (value is string)
                        {
                            Writer.WriteKeyword(key);
                            Writer.WriteText((string) value);
                        }
                        else if (value is int)
                            Writer.WriteKeyword(key + value);
                        else if (value is DateTime)
                        {
                            var dateTime = (DateTime) value;
                            Writer.WriteKeyword(key);
                            Writer.WriteKeyword("yr" + dateTime.Year);
                            Writer.WriteKeyword("mo" + dateTime.Month);
                            Writer.WriteKeyword("dy" + dateTime.Day);
                            Writer.WriteKeyword("hr" + dateTime.Hour);
                            Writer.WriteKeyword("min" + dateTime.Minute);
                            Writer.WriteKeyword("sec" + dateTime.Second);
                        }
                        else
                            Writer.WriteKeyword(key);

                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }
                
                // Write font table
                Writer.WriteStartGroup();
                Writer.WriteKeyword(Consts.Fonttbl);
                for (var count = 0; count < FontTable.Count; count ++)
                {
                    //string f = myFontTable[ count ] ;
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("f" + count);
                    var f = FontTable[count];
                    Writer.WriteText(f.Name);
                    if (f.Charset != 1)
                        Writer.WriteKeyword("fcharset" + f.Charset);
                    Writer.WriteEndGroup();
                }
                Writer.WriteEndGroup();

                // Write color table
                Writer.WriteStartGroup();
                Writer.WriteKeyword(Consts.Colortbl);
                Writer.WriteRaw(";");

                for (var count = 0; count < ColorTable.Count; count ++)
                {
                    var colorTable = ColorTable[count];
                    Writer.WriteKeyword("red" + colorTable.R);
                    Writer.WriteKeyword("green" + colorTable.G);
                    Writer.WriteKeyword("blue" + colorTable.B);
                    Writer.WriteRaw(";");
                }
                Writer.WriteEndGroup();

                // Write list table
                if (ListTable != null && ListTable.Count > 0)
                {
                    if (DebugMode)
                        Writer.WriteRaw(Environment.NewLine);
                    
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("listtable", true);
                    
                    foreach (var list in ListTable)
                    {
                        if (DebugMode)
                        {
                            Writer.WriteRaw(Environment.NewLine);
                        }
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("list");
                        Writer.WriteKeyword("listtemplateid" + list.ListTemplateId);
                        
                        if (list.ListHybrid)
                            Writer.WriteKeyword("listhybrid");
                        
                        if (DebugMode)
                            Writer.WriteRaw(Environment.NewLine);
                        
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("listlevel");
                        Writer.WriteKeyword("levelfollow" + list.LevelFollow);
                        Writer.WriteKeyword("leveljc" + list.LevelJc);
                        Writer.WriteKeyword("levelstartat" + list.LevelStartAt);
                        Writer.WriteKeyword("levelnfc" + Convert.ToInt32(list.LevelNfc));
                        Writer.WriteKeyword("levelnfcn" + Convert.ToInt32(list.LevelNfc));
                        Writer.WriteKeyword("leveljc" + list.LevelJc);

                        if (string.IsNullOrEmpty(list.LevelText) == false)
                        {
                            Writer.WriteStartGroup();
                            Writer.WriteKeyword("leveltext");
                            Writer.WriteKeyword("'0" + list.LevelText.Length);
                            if (list.LevelNfc == RtfLevelNumberType.Bullet)
                                Writer.WriteUnicodeText(list.LevelText);
                            else
                                Writer.WriteText(list.LevelText, false);

                            Writer.WriteEndGroup();
                            if (list.LevelNfc == RtfLevelNumberType.Bullet)
                            {
                                var f = FontTable["Wingdings"];
                                if (f != null)
                                    Writer.WriteKeyword("f" + f.Index);
                            }
                            else
                            {
                                Writer.WriteStartGroup();
                                Writer.WriteKeyword("levelnumbers");
                                Writer.WriteKeyword("'01");
                                Writer.WriteEndGroup();
                            }
                        }
                        Writer.WriteEndGroup();

                        Writer.WriteKeyword("listid" + list.ListId);
                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }

                // Write list overried table
                if (ListOverrideTable != null && ListOverrideTable.Count > 0)
                {
                    if (DebugMode)
                        Writer.WriteRaw(Environment.NewLine);
                    
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("listoverridetable");
                    foreach (var listOverride in ListOverrideTable)
                    {
                        if (DebugMode)
                        {
                            Writer.WriteRaw(Environment.NewLine);
                        }
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("listoverride");
                        Writer.WriteKeyword("listid" + listOverride.ListId);
                        Writer.WriteKeyword("listoverridecount" + listOverride.ListOverrideCount);
                        Writer.WriteKeyword("ls" + listOverride.Id);
                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }

                if (DebugMode)
                    Writer.WriteRaw(Environment.NewLine);
                
                Writer.WriteKeyword("viewkind1");
            }
        }

        /// <summary>
        /// Write the end of the document
        /// </summary>
        public void WriteEndDocument()
        {
            if (CollectionInfo == false)
                Writer.WriteEndGroup();
            
            Writer.Flush();
        }
        #endregion

        #region WriteStartEndHeader
        /// <summary>
        /// Write start from header
        /// </summary>
        public void WriteStartHeader()
        {
            if (CollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("header");
            }
        }

        /// <summary>
        /// Write end from header
        /// </summary>
        public void WriteEndHeader()
        {
            if (CollectionInfo == false)
            {
                Writer.WriteEndGroup();
            }
        }
        #endregion

        #region WriteStartEndFooter
        /// <summary>
        /// Write start from footer
        /// </summary>
        public void WriteStartFooter()
        {
            if (CollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("footer");
            }
        }

        /// <summary>
        /// Write end from footer
        /// </summary>
        public void WriteEndFooter()
        {
            if (CollectionInfo == false)
                Writer.WriteEndGroup();
        }
        #endregion

        #region WriteStartEndParagraph
        /// <summary>
        /// Write start from paragraph
        /// </summary>
        public void WriteStartParagraph()
        {
            WriteStartParagraph(new DocumentFormatInfo());
        }

        /// <summary>
        /// Write end of paragraph
        /// </summary>
        /// <param name="info">format</param>
        // ReSharper disable once MemberCanBePrivate.Global
        public void WriteStartParagraph(DocumentFormatInfo info)
        {
            if (CollectionInfo)
            {
                //myFontTable.Add("Wingdings");
            }
            else
            {
                if (_firstParagraph)
                {
                    _firstParagraph = false;
                    Writer.WriteRaw(Environment.NewLine);
                    //myWriter.WriteKeyword("par");
                }
                else
                {
                    Writer.WriteKeyword("par");
                }
                if (info.ListId >= 0)
                {
                    Writer.WriteKeyword("pard");
                    Writer.WriteKeyword("ls" + info.ListId);

                    if (_lastParagraphInfo != null)
                    {
                        if (_lastParagraphInfo.ListId >= 0)
                        {
                            Writer.WriteKeyword("pard");
                        }
                    }
                }

                switch (info.Align)
                {
                    case RtfAlignment.Left:
                        Writer.WriteKeyword("ql");
                        break;

                    case RtfAlignment.Center:
                        Writer.WriteKeyword("qc");
                        break;
                    
                    case RtfAlignment.Right:
                        Writer.WriteKeyword("qr");
                        break;
                    
                    case RtfAlignment.Justify:
                        Writer.WriteKeyword("qj");
                        break;
                }

                if (info.ParagraphFirstLineIndent != 0)
                {
                    Writer.WriteKeyword("fi" + Convert.ToInt32(
                        info.ParagraphFirstLineIndent*400/info.StandTabWidth));
                }
                else
                    Writer.WriteKeyword("fi0");

                if (info.LeftIndent != 0)
                {
                    Writer.WriteKeyword("li" + Convert.ToInt32(
                        info.LeftIndent*400/info.StandTabWidth));
                }
                else
                {
                    Writer.WriteKeyword("li0");
                }
                Writer.WriteKeyword("plain");
            }
            _lastParagraphInfo = info;
        }

        /// <summary>
        /// end write paragraph
        /// </summary>
        public void WriteEndParagraph()
        {
        }
        #endregion

        #region WriteText
        /// <summary>
        /// Write plain text
        /// </summary>
        /// <param name="text">text</param>
        public void WriteText(string text)
        {
            if (text != null && CollectionInfo == false)
                Writer.WriteText(text);
        }
        #endregion

        #region WriteFont
        /// <summary>
        /// Write font format
        /// </summary>
        /// <param name="font">font</param>
        public void WriteFont(System.Drawing.Font font)
        {
            if (font == null)
                throw new ArgumentNullException("font");
            if (CollectionInfo)
                FontTable.Add(font.Name);
            else
            {
                var index = FontTable.IndexOf(font.Name);
                
                if (index >= 0)
                    Writer.WriteKeyword("f" + index);
                
                if (font.Bold)
                    Writer.WriteKeyword("b");
                
                if (font.Italic)
                    Writer.WriteKeyword("i");
                
                if (font.Underline)
                    Writer.WriteKeyword("ul");
                
                if (font.Strikeout)
                    Writer.WriteKeyword("strike");
                
                Writer.WriteKeyword("fs" + Convert.ToInt32(font.Size*2));
            }
        }
        #endregion

        #region WriteString
        /// <summary>
        /// Start write of formatted text
        /// </summary>
        /// <param name="info">format</param>
        /// <remarks>
        /// This function must assort with WriteEndString
        /// </remarks>
        // ReSharper disable once MemberCanBePrivate.Global
        public void WriteStartString(DocumentFormatInfo info)
        {
            if (CollectionInfo)
            {
                FontTable.Add(info.FontName);
                ColorTable.Add(info.TextColor);
                ColorTable.Add(info.BackColor);
                if (info.BorderColor.A != 0)
                {
                    ColorTable.Add(info.BorderColor);
                }
                return;
            }

            if (!string.IsNullOrEmpty(info.Link))
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("field");
                Writer.WriteStartGroup();
                Writer.WriteKeyword("fldinst", true);
                Writer.WriteStartGroup();
                Writer.WriteKeyword("hich");
                Writer.WriteText(" HYPERLINK \"" + info.Link + "\"");
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
                Writer.WriteStartGroup();
                Writer.WriteKeyword("fldrslt");
                Writer.WriteStartGroup();
            }

            switch (info.Align)
            {
                case RtfAlignment.Left:
                    Writer.WriteKeyword("ql");
                    break;
            
                case RtfAlignment.Center:
                    Writer.WriteKeyword("qc");
                    break;
                
                case RtfAlignment.Right:
                    Writer.WriteKeyword("qr");
                    break;
               
                case RtfAlignment.Justify:
                    Writer.WriteKeyword("qj");
                    break;
            }

            Writer.WriteKeyword("plain");

            int index = FontTable.IndexOf(info.FontName);
            
            if (index >= 0)
                Writer.WriteKeyword("f" + index);
            
            if (info.Bold)
                Writer.WriteKeyword("b");
            
            if (info.Italic)
                Writer.WriteKeyword("i");
            
            if (info.Underline)
                Writer.WriteKeyword("ul");
            
            if (info.Strikeout)
                Writer.WriteKeyword("strike");
            
            Writer.WriteKeyword("fs" + Convert.ToInt32(info.FontSize*2));

            // Back color
            index = ColorTable.IndexOf(info.BackColor);
            if (index >= 0)
                Writer.WriteKeyword("chcbpat" + Convert.ToString(index + 1));

            index = ColorTable.IndexOf(info.TextColor);
            
            if (index >= 0)
                Writer.WriteKeyword("cf" + Convert.ToString(index + 1));
            
            if (info.Subscript)
                Writer.WriteKeyword("sub");
            
            if (info.Superscript)
                Writer.WriteKeyword("super");
            
            if (info.NoWwrap)
                Writer.WriteKeyword("nowwrap");
            
            if (info.LeftBorder
                || info.TopBorder
                || info.RightBorder
                || info.BottomBorder)
            {
                // Border color
                if (info.BorderColor.A != 0)
                {
                    Writer.WriteKeyword("chbrdr");
                    Writer.WriteKeyword("brdrs");
                    Writer.WriteKeyword("brdrw10");
                    index = ColorTable.IndexOf(info.BorderColor);
                    if (index >= 0)
                    {
                        Writer.WriteKeyword("brdrcf" + Convert.ToString(index + 1));
                    }
                }
            }
        }

        // ReSharper disable once MemberCanBePrivate.Global
        /// <summary>
        /// End writing of a formatted string. WriteStartString and WriteString have to be called before this method
        /// </summary>
        /// <param name="info"></param>
        public void WriteEndString(DocumentFormatInfo info)
        {
            if (CollectionInfo)
            {
                return;
            }

            if (info.Subscript)
                Writer.WriteKeyword("sub0");
            if (info.Superscript)
                Writer.WriteKeyword("super0");

            if (info.Bold)
                Writer.WriteKeyword("b0");
            
            if (info.Italic)
                Writer.WriteKeyword("i0");
            
            if (info.Underline)
                Writer.WriteKeyword("ul0");
            
            if (info.Strikeout)
                Writer.WriteKeyword("strike0");

            if (!string.IsNullOrEmpty(info.Link))
            {
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// Write formatted string. Call WriteStartString before this method
        /// </summary>
        /// <param name="text">text</param>
        /// <param name="info">format</param>
        public void WriteString(string text, DocumentFormatInfo info)
        {
            if (CollectionInfo)
            {
                FontTable.Add(info.FontName);
                ColorTable.Add(info.TextColor);
                ColorTable.Add(info.BackColor);
            }
            else
            {
                WriteStartString(info);

                if (info.Multiline)
                {
                    if (text != null)
                    {
                        text = text.Replace("\n", "");
                        var reader = new StringReader(text);
                        var strLine = reader.ReadLine();
                        var count = 0;

                        while (strLine != null)
                        {
                            if (count > 0)
                                Writer.WriteKeyword("line");

                            count ++;
                            Writer.WriteText(strLine);
                            strLine = reader.ReadLine();
                        }
                        
                        reader.Close();
                    }
                }
                else
                    Writer.WriteText(text);

                WriteEndString(info);
            }
        }

        /// <summary>
        /// End write string
        /// </summary>
        public void WriteEndString()
        {
        }
        #endregion

        #region WriteBookmark
        /// <summary>
        /// Start write of bookmark
        /// </summary>
        /// <param name="strName">bookmark name</param>
        public void WriteStartBookmark(string strName)
        {
            if (CollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("bkmkstart", true);
                Writer.WriteKeyword("f0");
                Writer.WriteText(strName);
                Writer.WriteEndGroup();

                Writer.WriteStartGroup();
                Writer.WriteKeyword("bkmkend", true);
                Writer.WriteKeyword("f0");
                Writer.WriteText(strName);
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// End write of bookmark
        /// </summary>
        /// <param name="strName">bookmark name</param>
        public void WriteEndBookmark(string strName)
        {
        }
        #endregion

        #region WriteLineBreak
        /// <summary>
        /// Write a line break
        /// </summary>
        public void WriteLineBreak()
        {
            if (CollectionInfo == false)
            {
                Writer.WriteKeyword("line");
            }
        }
        #endregion

        #region WriteImage
        /// <summary>
        /// Write image
        /// </summary>
        /// <param name="image">Image</param>
        /// <param name="width">Pixel width</param>
        /// <param name="height">Pixel height</param>
        /// <param name="imageData">Image binary data</param>
        public void WriteImage(Image image, int width, int height, byte[] imageData)
        {
            if (imageData == null)
                return;

            var memoryStream = new MemoryStream();
            image.Save(memoryStream, ImageFormat.Jpeg);
            memoryStream.Close();
            var bs = memoryStream.ToArray();
            Writer.WriteStartGroup();

            Writer.WriteKeyword("pict");
            Writer.WriteKeyword("jpegblip");
            Writer.WriteKeyword("picscalex" + Convert.ToInt32(width*100.0/image.Size.Width));
            Writer.WriteKeyword("picscaley" + Convert.ToInt32(height*100.0/image.Size.Height));
            Writer.WriteKeyword("picwgoal" + Convert.ToString(image.Size.Width*15));
            Writer.WriteKeyword("pichgoal" + Convert.ToString(image.Size.Height*15));
            Writer.WriteBytes(bs);
            Writer.WriteEndGroup();
        }
        #endregion
    }
}