namespace MsgReader.Rtf
{
    /// <summary>
    /// Dom table cell 
    /// </summary>
    internal class DomTableCell : DomElement
    {
        #region Properties
        /// <summary>
        /// Row span
        /// </summary>
        public int RowSpan { get; set; }

        /// <summary>
        /// Col span
        /// </summary>
        public int ColSpan { get; set; }

        /// <summary>
        /// Left padding
        /// </summary>
        public int PaddingLeft { get; set; }

        /// <summary>
        /// Left padding in fact
        /// </summary>
        public int RuntimePaddingLeft
        {
            get
            {
                if (PaddingLeft != int.MinValue)
                    return PaddingLeft;

                if (Parent != null)
                {
                    var parent = ((DomTableRow) Parent).PaddingLeft;
                    if (parent != int.MinValue)
                        return parent;
                }
                return 0;
            }
        }

        /// <summary>
        /// Top padding
        /// </summary>
        public int PaddingTop { get; set; }

        /// <summary>
        /// Top padding in fact
        /// </summary>
        public int RuntimePaddingTop
        {
            get
            {
                if (PaddingTop != int.MinValue)
                    return PaddingTop;

                if (Parent != null)
                {
                    var parent = ((DomTableRow)Parent).PaddingTop;
                    if (parent != int.MinValue)
                        return parent;
                }
                return 0;
            }
        }

        /// <summary>
        /// Right padding
        /// </summary>
        public int PaddingRight { get; set; }

        /// <summary>
        /// Right padding in fact
        /// </summary>
        public int RuntimePaddingRight
        {
            get
            {
                if (PaddingRight != int.MinValue)
                {
                    return PaddingRight;
                }
                if (Parent != null)
                {
                    var parent = ((DomTableRow) Parent).PaddingRight;
                    if (parent != int.MinValue)
                    {
                        return parent;
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// Bottom padding
        /// </summary>
        public int PaddingBottom { get; set; }

        /// <summary>
        /// Bottom padding in fact
        /// </summary>
        public int RuntimePaddingBottom
        {
            get
            {
                if (PaddingBottom != int.MinValue)
                {
                    return PaddingBottom;
                }
                if (Parent != null)
                {
                    var p = ((DomTableRow) Parent).PaddingBottom;
                    if (p != int.MinValue)
                    {
                        return p;
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// Vertial alignment
        /// </summary>
        public RtfVerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }

        /// <summary>
        /// Allow multiline
        /// </summary>
        public bool Multiline { get; set; }

        /// <summary>
        /// Left position
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Cell merged by another cell which this property specified
        /// </summary>
        public DomTableCell OverrideCell { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        public DomTableCell()
        {
            Multiline = true;
            Format = new DocumentFormatInfo();
            VerticalAlignment = RtfVerticalAlignment.Top;
            PaddingBottom = int.MinValue;
            PaddingRight = int.MinValue;
            PaddingTop = int.MinValue;
            PaddingLeft = int.MinValue;
            ColSpan = 1;
            RowSpan = 1;
            Height = 0;
            Width = 0;
            Left = 0;
            Format.BorderWidth = 1;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            if (OverrideCell == null)
            {
                if (RowSpan != 1 || ColSpan != 1)
                {
                    return "Cell: RowSpan:" + RowSpan + " ColSpan:" + ColSpan + " Width:" + Width;
                }
                return "Cell:Width:" + Width;
            }
            return "Cell:Overrided";
        }
        #endregion
    }
}