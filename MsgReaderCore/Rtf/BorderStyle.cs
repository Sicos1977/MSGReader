//
// BorderStyle.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
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