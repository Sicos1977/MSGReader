//
// DomElementList.cs
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
    /// RTF element's list
    /// </summary>
    internal class DomElementList : CollectionBase
    {
        #region Properties
        /// <summary>
        /// Get the last element in the list
        /// </summary>
        public DomElement LastElement => Count > 0 ? (DomElement)List[Count - 1] : null;

        /// <summary>
        /// Get the element at special index
        /// </summary>
        /// <param name="index">index</param>
        /// <returns>element</returns>
        public DomElement this[int index] => (DomElement) List[index];
        #endregion

        #region Add
        /// <summary>
        /// Add element
        /// </summary>
        /// <param name="element">element</param>
        /// <returns>index</returns>
        public int Add(DomElement element)
        {
            return List.Add(element);
        }
        #endregion

        #region Insert
        /// <summary>
        /// Insert element
        /// </summary>
        /// <param name="index">special index</param>
        /// <param name="element">element</param>
        public void Insert(int index, DomElement element)
        {
            List.Insert(index, element);
        }
        #endregion

        #region IndexOf
        /// <summary>
        /// Get the index of special element that starts with 0.
        /// </summary>
        /// <param name="element">element</param>
        /// <returns>index , if not find element , then return -1</returns>
        public int IndexOf(DomElement element)
        {
            return List.IndexOf(element);
        }
        #endregion

        #region Remove
        /// <summary>
        /// Remove element
        /// </summary>
        /// <param name="node">element</param>
        public void Remove(DomElement node)
        {
            List.Remove(node);
        }
        #endregion

        #region ToArray
        /// <summary>
        /// Return element array
        /// </summary>
        /// <returns>array</returns>
        public DomElement[] ToArray()
        {
            return (DomElement[]) InnerList.ToArray(typeof (DomElement));
        }
        #endregion
    }
}