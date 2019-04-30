//
// DomTablRow.cs
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

using System.Collections;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Table row
    /// </summary>
    internal class DomTableRow : DomElement
    {
        #region Properties
        /// <summary>
        /// Cell settings
        /// </summary>
        internal ArrayList CellSettings { get; set; }

        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }

        /// <summary>
        /// Document level
        /// </summary>
        public int Level { get; set; }

        /// <summary>
        /// row index
        /// </summary>
        internal int RowIndex { get; set; }

        /// <summary>
        /// Is the last row
        /// </summary>
        public bool IsLastRow { get; set; }

        /// <summary>
        /// Is header row
        /// </summary>
        public bool Header { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }
        
        /// <summary>
        /// Padding left
        /// </summary>
        public int PaddingLeft { get; set; }

        /// <summary>
        /// Top padding
        /// </summary>
        public int PaddingTop { get; set; }


        /// <summary>
        /// Right padding
        /// </summary>
        public int PaddingRight { get; set; }

        /// <summary>
        /// Bottom padding
        /// </summary>
        public int PaddingBottom { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }
        #endregion

        #region Constructor
        public DomTableRow()
        {
            Width = 0;
            PaddingBottom = int.MinValue;
            PaddingRight = int.MinValue;
            PaddingTop = int.MinValue;
            PaddingLeft = int.MinValue;
            Height = 0;
            Header = false;
            IsLastRow = false;
            RowIndex = 0;
            Level = 1;
            Format = new DocumentFormatInfo();
            CellSettings = new ArrayList();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Row " + RowIndex;
        }
        #endregion
    }
}
