//
// DomShape.cs
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
    /// Rtf shape element
    /// </summary>
    internal class DomShape : DomElement
    {
        #region Properties
        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable UnusedAutoPropertyAccessor.Global
        /// <summary>
        /// Left position
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Top position
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Z index
        /// </summary>
        public int ZIndex { get; set; }

        /// <summary>
        /// Shape id
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Ext attribute
        /// </summary>
        public StringAttributeCollection ExtAttrbutes { get; set; }
        // ReSharper restore UnusedAutoPropertyAccessor.Global
        // ReSharper restore MemberCanBePrivate.Global
        #endregion

        #region Constructor
        public DomShape()
        {
            ExtAttrbutes = new StringAttributeCollection();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Shape:Left:" + Left + " Top:" + Top + " Width:" + Width + " Height:" + Height;
        }
        #endregion
    }
}