//
// DomParagraph.cs
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

using System;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf paragraph element
    /// </summary>
    internal class DomParagraph : DomElement
    {
        #region Properties
        /// <summary>
        /// If it is generated from a template
        /// </summary>
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool IsTemplateGenerated { get; internal set; }

        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }
        #endregion

        #region Constructor
        public DomParagraph()
        {
            Format = new DocumentFormatInfo();
            IsTemplateGenerated = false;
        }
        #endregion

        #region InnerText
        public override string InnerText => base.InnerText + Environment.NewLine;
        #endregion

        #region ToString
        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("Paragraph");
            if (Format != null)
            {
                stringBuilder.Append("(" + Format.Align + ")");
                if (Format.ListId >= 0)
                    stringBuilder.Append("ListID:" + Format.ListId);
            }

            return stringBuilder.ToString();
        }
        #endregion
    }
}