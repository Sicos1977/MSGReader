//
// DomTableCell.cs
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