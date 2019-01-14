//
// DomImage.cs
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

namespace MsgReader.Rtf
{
    /// <summary>
    /// Image element
    /// </summary>
    internal class DomImage : DomElement
    {
        #region Properties
        /// <summary>
        /// Id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Data
        /// </summary>
        public byte[] Data { get; set; }

        /// <summary>
        /// Data as base64 format
        /// </summary>
        public string Base64Data
        {
            get
            {
                return Data == null ? null : Convert.ToBase64String(Data);
            }
            set
            {
                Data = !string.IsNullOrEmpty(value) ? Convert.FromBase64String(value) : null;
            }
        }

        /// <summary>
        /// Scale rate at the X coordinate, in percent unit.
        /// </summary>
        public int ScaleX { get; set; }

        /// <summary>
        /// Scale rate at the Y coordinate , in percent unit.
        /// </summary>
        public int ScaleY { get; set; }

        /// <summary>
        /// Desired width
        /// </summary>
        public int DesiredWidth { get; set; }

        /// <summary>
        /// Desired height
        /// </summary>
        public int DesiredHeight { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Picture type
        /// </summary>
        public RtfPictureType PicType { get; set; }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }
        #endregion

        #region Constructor
        public DomImage()
        {
            Format = new DocumentFormatInfo();
            ScaleY = 100;
            ScaleX = 100;
            PicType = RtfPictureType.Jpegblip;
            Height = 0;
            Width = 0;
            DesiredHeight = 0;
            DesiredWidth = 0;
            Data = null;
            Id = null;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var text = "Image:" + Width + "*" + Height;
            if (Data != null && Data.Length > 0)
                text = text + " " + Convert.ToDouble( Data.Length / 1024.0).ToString("0.00") + "KB";
            return text;
        }
        #endregion
    }
}
