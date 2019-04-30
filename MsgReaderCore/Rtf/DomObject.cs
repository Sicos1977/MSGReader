//
// DomObject.cs
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
using System.Collections.Generic;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF Dom object
    /// </summary>
    internal class DomObject : DomElement
    {
        #region Fields
        private Dictionary<string, string> _customAttributes = new Dictionary<string, string>();
        #endregion

        #region Properties
        /// <summary>
        /// Custom attributes
        /// </summary>
        public Dictionary<string, string> CustomAttributes
        {
            get { return _customAttributes ?? (_customAttributes = new Dictionary<string, string>()); }
            set
            {
                _customAttributes = value;
            }
        }

        /// <summary>
        /// Type
        /// </summary>
        public RtfObjectType Type { get; set; }

        /// <summary>
        /// Class name
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Content as byte array
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// Content type
        /// </summary>
        public string ContentText
        {
            get
            {
                if (Content == null || Content.Length == 0)
                    return null;

                return System.Text.Encoding.Default.GetString(Content);
            }
        }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Scaling X factor
        /// </summary>
        public int ScaleX { get; set; }

        /// <summary>
        /// Scaling Y factor
        /// </summary>
        public int ScaleY { get; set; }

        /// <summary>
        /// Result
        /// </summary>
        public ElementContainer Result
        {
            get
            {
                foreach (DomElement element in Elements)
                {
                    if (element is ElementContainer)
                    {
                        var container = (ElementContainer)element;
                        if (container.Name == Consts.Result)
                            return container;
                    }
                }
                return null;
            }
        }

        #endregion

        #region Constructor
        public DomObject()
        {
            ScaleX = 100;
            ScaleY = 100;
            Height = 0;
            Width = 0;
            Content = null;
            Name = null;
            ClassName = null;
            Type = RtfObjectType.Emb;
        }
        #endregion
        
        #region ToString
        public override string ToString()
        {
            var text = "Object:" + Width + "*" + Height;
            if (Content != null && Content.Length > 0)
                text = text + " " + Convert.ToDouble(Content.Length / 1024.0).ToString("0.00") + "KB";
            return text;
        }
        #endregion
    }
}
