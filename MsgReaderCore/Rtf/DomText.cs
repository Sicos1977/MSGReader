//
// DomText.cs
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

using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf text element
    /// </summary>
    internal class DomText : DomElement
    {
        #region Properties
        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }

        /// <summary>
        /// Text
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Inner text
        /// </summary>
        public override string InnerText => Text;
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        public DomText()
        {
            Format = new DocumentFormatInfo();
            // Text element can not contains any child element
            Locked = true;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("Text");
            if (Format != null)
            {
                if (Format.Hidden)
                    stringBuilder.Append("(Hidden)");
            }
            stringBuilder.Append(":" + Text);
            return stringBuilder.ToString();
        }
        #endregion
    }
}