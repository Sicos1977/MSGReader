//
// DomElement.cs
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
    /// RTF dom element (this is the most base element type)
    /// </summary>
    internal abstract class DomElement
    {
        #region Fields
        /// <summary>
        /// Native level in RTF document
        /// </summary>
        internal int NativeLevel = -1;

        private DomDocument _ownerDocument;
        #endregion

        #region Properties
        /// <summary>
        /// RTF native attribute
        /// </summary>
        public AttributeList Attributes { get; set; }

        /// <summary>
        /// child elements list
        /// </summary>
        public DomElementList Elements { get; }

        /// <summary>
        /// The document which owned this element
        /// </summary>
        public DomDocument OwnerDocument
        {
            get { return _ownerDocument; }
            set
            {
                _ownerDocument = value;
                foreach (DomElement element in Elements)
                {
                    element.OwnerDocument = value;
                }
            }
        }

        /// <summary>
        /// Parent element
        /// </summary>
        public DomElement Parent { get; private set; }

        public virtual string InnerText
        {
            get
            {
                var stringBuilder = new StringBuilder();
                if (Elements != null)
                {
                    foreach (DomElement element in Elements)
                    {
                        stringBuilder.Append(element.InnerText);
                    }
                }
                return stringBuilder.ToString();
            }
        }


        /// <summary>
        /// Whether element is locked , if element is lock , it can not append chidl element
        /// </summary>
        public bool Locked { get; set; }
        #endregion

        #region Protected constructor
        protected DomElement()
        {
            Elements = new DomElementList();
            Attributes = new AttributeList();
        }
        #endregion

        #region HasAttribute
        public bool HasAttribute(string name)
        {
            return Attributes.Contains(name);
        }
        #endregion

        #region GetAttributeValue
        public int GetAttributeValue(string name, int defaultValue)
        {
            return Attributes.Contains(name) ? Attributes[name] : defaultValue;
        }
        #endregion

        #region AppendChild
        /// <summary>
        /// Append child element
        /// </summary>
        /// <param name="element">child element</param>
        /// <returns>index of element</returns>
        public int AppendChild(DomElement element)
        {
            CheckLocked();
            element.Parent = this;
            element.OwnerDocument = _ownerDocument;
            return Elements.Add(element);
        }
        #endregion

        #region SetAttribute
        /// <summary>
        /// Set attribute
        /// </summary>
        /// <param name="name">name</param>
        /// <param name="value">value</param>
        public void SetAttribute(string name, int value)
        {
            CheckLocked();
            Attributes[name] = value;
        }
        #endregion

        #region CheckLocked
        private void CheckLocked()
        {
            if (Locked)
            {
                throw new InvalidOperationException("Element locked");
            }
        }
        #endregion

        #region SetLockedDeeply
        public void SetLockedDeeply(bool locked)
        {
            Locked = locked;
            if (Elements != null)
            {
                foreach (DomElement element in Elements)
                {
                    element.SetLockedDeeply(locked);
                }
            }
        }
        #endregion

        #region ToDomString
        public virtual string ToDomString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(ToString());
            ToDomString(Elements, stringBuilder, 1);
            return stringBuilder.ToString();
        }

        protected void ToDomString(DomElementList elements, StringBuilder stringBuilder, int level)
        {
            foreach (DomElement element in elements)
            {
                stringBuilder.Append(Environment.NewLine);
                stringBuilder.Append(new string(' ', level*4));
                stringBuilder.Append(element);
                ToDomString(element.Elements, stringBuilder, level + 1);
            }
        }
        #endregion
    }
}