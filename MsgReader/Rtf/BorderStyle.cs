using System.Drawing;
using System.Drawing.Drawing2D;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf boder style
    /// </summary>
    internal class BorderStyle
    {
        #region Properties
        public bool Left { get; set; }
        public bool Top { get; set; }
        public bool Right { get; set; }
        public bool Bottom { get; set; }
        public DashStyle Style { get; set; }
        public Color Color { get; set; }
        public bool Thickness { get; set; }
        #endregion

        #region Constructor
        public BorderStyle()
        {
            Color = Color.Black;
            Style = DashStyle.Solid;
        }
        #endregion

        #region EqualsValue
        public bool EqualsValue(BorderStyle borderStyle)
        {
            if (borderStyle == this)
                return true;
            
            if (borderStyle == null)
                return false;

            return borderStyle.Bottom == Bottom && 
                   borderStyle.Color == Color && 
                   borderStyle.Left == Left &&
                   borderStyle.Right == Right && 
                   borderStyle.Style == Style && 
                   borderStyle.Top == Top &&
                   borderStyle.Thickness == Thickness;
        }
        #endregion
    }
}