//
// StringAttribute.cs
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
    internal class StringAttribute
    {
        #region Properties
        /// <summary>
        /// Name of the attribute
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Value of the attribute
        /// </summary>
        public string Value { get; set; }
        #endregion

        #region Constructor
        public StringAttribute()
        {
            Value = null;
            Name = null;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return Name + "=" + Value;
        }
        #endregion
    }

    internal class StringAttributeCollection : System.Collections.CollectionBase
    {
        #region Properties
        public string this[string name]
        {
            get
            {
                foreach (StringAttribute stringAttribute in this)
                {
                    if (stringAttribute.Name == name)
                        return stringAttribute.Value;
                }
                return null;
            }
            set
            {
                foreach (StringAttribute stringAttribute in this)
                {
                    if (stringAttribute.Name == name)
                    {
                        if (value == null)
                            List.Remove(stringAttribute);
                        else
                            stringAttribute.Value = value;
                        return;
                    }
                }
                if (value != null)
                {
                    var stringAttribute = new StringAttribute { Name = name, Value = value };
                    List.Add(stringAttribute);
                }
            }
        }
        #endregion

        #region Add
        /// <summary>
        /// Add's an item to the list and returns the index
        /// </summary>
        /// <param name="stringAttribute"></param>
        /// <returns></returns>
        public int Add(StringAttribute stringAttribute)
        {
            return List.Add(stringAttribute);
        }
        #endregion

        #region Remove
        /// <summary>
        /// Removes an item from the list
        /// </summary>
        /// <param name="stringAttribute"></param>
        public void Remove(StringAttribute stringAttribute)
        {
            List.Remove(stringAttribute);
        }
        #endregion
    }
}
