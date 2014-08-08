using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF Document format information
    /// </summary>
    internal class DocumentFormatInfo
    {
        #region Fields
        internal bool ReadText = true;
        private bool  _subscript;
        private bool _superscript;
        #endregion

        #region Properties
        /// <summary>
        /// If this instance is created by Clone , return the parent instance
        /// </summary>
        public DocumentFormatInfo Parent { get; private set; }

        /// <summary>
        /// Display left border line
        /// </summary>
        public bool LeftBorder { get; set; }

        /// <summary>
        /// Display top border line
        /// </summary>
        public bool TopBorder { get; set; }

        /// <summary>
        /// Display right border line
        /// </summary>
        public bool RightBorder { get; set; }

        /// <summary>
        /// Display bottom border line
        /// </summary>
        public bool BottomBorder { get; set; }

        /// <summary>
        /// Border line color
        /// </summary>
        public Color BorderColor { get; set; }

        /// <summary>
        /// Border line width
        /// </summary>
        public int BorderWidth { get; set; }

        /// <summary>
        /// Border style
        /// </summary>
        public DashStyle BorderStyle { get; set; }

        /// <summary>
        /// Border thicknes
        /// </summary>
        public bool BorderThickness { get; set; }

        /// <summary>
        /// Border spacing
        /// </summary>
        public int BorderSpacing { get; set; }

        /// <summary>
        /// Word wrap
        /// </summary>
        public bool Multiline { get; set; }

        /// <summary>
        /// Standard tab width
        /// </summary>
        public int StandTabWidth { get; set; }

        /// <summary>
        /// indent of first line in a paragraph
        /// </summary>
        [DefaultValue(0)]
        public int ParagraphFirstLineIndent { get; set; }

        /// <summary>
        /// Indent of wholly paragraph
        /// </summary>
        public int LeftIndent { get; set; }

        /// <summary>
        /// character spacing
        /// </summary>
        public int Spacing { get; set; }

        /// <summary>
        /// line spacing
        /// </summary>
        public int LineSpacing { get; set; }

        /// <summary>
        /// Current line spacing is multiple extractly line spacing.
        /// </summary>
        public bool MultipleLineSpacing { get; set; }

        /// <summary>
        /// Spacing before paragrah
        /// </summary>
        public int SpacingBefore { get; set; }

        /// <summary>
        /// Spacing after paragraph
        /// </summary>
        public int SpacingAfter { get; set; }

        /// <summary>
        /// Text alignment
        /// </summary>
        public RtfAlignment Align { get; set; }

        /// <summary>
        /// Page break
        /// </summary>
        public bool PageBreak { get; set; }

        /// <summary>
        /// Nest level in native rtf document
        /// </summary>
        public int NativeLevel { get; set; }

        public System.Drawing.Font Font
        {
            set
            {
                if (value == null) return;
                FontName = value.Name;
                FontSize = value.Size;
                Bold = value.Bold;
                Italic = value.Italic;
                Underline = value.Underline;
                Strikeout = value.Strikeout;
            }
        }

        /// <summary>
        /// Font name
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        public float FontSize { get; set; }
        
        /// <summary>
        /// Bold style
        /// </summary>
        public bool Bold { get; set; }

        /// <summary>
        /// Italic style
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// Underline style
        /// </summary>
        public bool Underline { get; set; }

        /// <summary>
        /// Strickout style
        /// </summary>
        public bool Strikeout { get; set; }

        /// <summary>
        /// Hidden text
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Text color
        /// </summary>
        public Color TextColor { get; set; }

        /// <summary>
        /// Back color
        /// </summary>
        public Color BackColor { get; set; }

        /// <summary>
        /// Link
        /// </summary>
        public string Link { get; set; }

        /// <summary>
        /// Superscript
        /// </summary>
        public bool Superscript
        {
            get
            {
                return _superscript;
            }
            set
            {
                _superscript = value;
                if (_superscript)
                    _subscript = false;
            }
        }

        /// <summary>
        /// Subscript
        /// </summary>
        public bool Subscript
        {
            get
            {
                return _subscript;
            }
            set
            {
                _subscript = value;
                if (_subscript)
                {
                    _superscript = false;
                }
            }
        }

        /// <summary>
        /// List override id 
        /// </summary>
        public int ListId { get; set; }

        /// <summary>
        /// No wrap in word
        /// </summary>
        public bool NoWwrap { get; set; }
        #endregion

        #region SetAlign
        public void SetAlign(StringAlignment align)
        {
            switch (align)
            {
                case StringAlignment.Center:
                    Align = RtfAlignment.Center;
                    break;

                case StringAlignment.Far:
                    Align = RtfAlignment.Right;
                    break;

                default:
                    Align = RtfAlignment.Left;
                    break;
            }
        }
        #endregion

        #region DocumentFormatInfo
        public DocumentFormatInfo()
        {
            NoWwrap = true;
            ListId = -1;
            Link = null;
            BackColor = Color.Empty;
            TextColor = Color.Black;
            Hidden = false;
            Strikeout = false;
            Underline = false;
            Italic = false;
            Bold = false;
            FontSize = 12f;
            FontName = System.Windows.Forms.Control.DefaultFont.Name;
            Align = RtfAlignment.Left;
            SpacingAfter = 0;
            SpacingBefore = 0;
            MultipleLineSpacing = false;
            LineSpacing = 0;
            Spacing = 0;
            LeftIndent = 0;
            ParagraphFirstLineIndent = 0;
            PageBreak = false;
            StandTabWidth = 100;
            BorderColor = Color.Black;
            Multiline = false;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderStyle = DashStyle.Solid;
            BorderWidth = 0;
            BottomBorder = false;
            RightBorder = false;
            TopBorder = false;
            LeftBorder = false;
            Parent = null;
        }
        #endregion

        #region EqualsSettings
        public bool EqualsSettings(DocumentFormatInfo format)
        {
            if (format == this)
                return true;
            if (format == null)
                return false;
            if (Align != format.Align)
                return false;
            if (BackColor != format.BackColor)
                return false;
            if (Bold != format.Bold)
                return false;
            if (BorderColor != format.BorderColor)
                return false;
            if (LeftBorder != format.LeftBorder)
                return false;
            if (TopBorder != format.TopBorder)
                return false;
            if (RightBorder != format.RightBorder)
                return false;
            if (BottomBorder != format.BottomBorder)
                return false;
            if (BorderStyle != format.BorderStyle)
                return false;
            if (BorderThickness != format.BorderThickness)
                return false;
            if (BorderSpacing != format.BorderSpacing)
                return false;
            if (ListId != format.ListId)
                return false;
            if (FontName != format.FontName)
                return false;
            // ReSharper disable once CompareOfFloatsByEqualityOperator
            if (FontSize != format.FontSize)
                return false;
            if (Italic != format.Italic)
                return false;
            if (Hidden != format.Hidden)
                return false;
            if (LeftIndent != format.LeftIndent)
                return false;
            if (LineSpacing != format.LineSpacing)
                return false;
            if (Link != format.Link)
                return false;
            if (Multiline != format.Multiline)
                return false;
            if (NoWwrap != format.NoWwrap)
                return false;
            if (ParagraphFirstLineIndent != format.ParagraphFirstLineIndent)
                return false;
            if (Spacing != format.Spacing)
                return false;
            if (StandTabWidth != format.StandTabWidth)
                return false;
            if (Strikeout != format.Strikeout)
                return false;
            if (Subscript != format.Subscript)
                return false;
            if (Superscript != format.Superscript)
                return false;
            if (TextColor != format.TextColor)
                return false;
            if (Underline != format.Underline)
                return false;
            if (ReadText != format.ReadText)
                return false;

            return true;
        }
        #endregion

        #region Clone
        /// <summary>
        /// Clone instance
        /// </summary>
        /// <returns>new instance</returns>
        public DocumentFormatInfo Clone()
        {
            return (DocumentFormatInfo) MemberwiseClone();
        }
        #endregion

        #region ResetText
        public void ResetText()
        {
            FontName = System.Windows.Forms.Control.DefaultFont.Name;
            Bold = false;
            Italic = false;
            Underline = false;
            Strikeout = false;
            TextColor = Color.Black;
            BackColor = Color.Empty;
            //Link = null ;
            Subscript = false;
            Superscript = false;
            Multiline = true;
            Hidden = false;
            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderStyle = DashStyle.Solid;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black ;
        }
        #endregion

        #region ResetParagraph
        public void ResetParagraph()
        {
            ParagraphFirstLineIndent = 0;
            Align = 0;
            ListId = -1;
            LeftIndent = 0;
            LineSpacing = 0;
            PageBreak = false;
            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderStyle = DashStyle.Solid;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black  ;
            MultipleLineSpacing = false;
            SpacingBefore = 0;
            SpacingAfter = 0;
        }
        #endregion

        #region Reset
        public void Reset()
        {
            ParagraphFirstLineIndent = 0;
            LeftIndent = 0;
            LeftIndent = 0;
            Spacing = 0;
            LineSpacing = 0;
            MultipleLineSpacing = false;
            SpacingBefore = 0;
            SpacingAfter = 0;
            Align = 0;
            FontName = System.Windows.Forms.Control.DefaultFont.Name;
            FontSize = 12;
            Bold = false;
            Italic = false;
            Underline = false;
            Strikeout = false;
            TextColor = Color.Black;
            BackColor = Color.Empty;
            Link = null;
            Subscript = false;
            Superscript = false;
            ListId = -1;
            Multiline = true;
            NoWwrap = true;
            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderStyle = DashStyle.Solid;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black;
            ReadText = true;
            NativeLevel = 0;
            Hidden = false;
        }
        #endregion
    }
}