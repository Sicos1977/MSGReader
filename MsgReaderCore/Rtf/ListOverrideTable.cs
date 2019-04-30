//
// ListOverrideTable.cs
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

using System.Collections.Generic;

namespace MsgReader.Rtf
{
    internal class ListOverrideTable : List<ListOverride>
    {
        #region GetById
        public ListOverride GetById(int id)
        {
            foreach (var listOverride in this)
            {
                if (listOverride.Id == id)
                    return listOverride;
            }
            return null;
        }
        #endregion
    }

    internal class ListOverride
    {
        #region Properties
        public int ListId { get; set; }

        /// <summary>
        /// List override count
        /// </summary>
        public int ListOverrideCount { get; set; }

        /// <summary>
        /// Internal Id
        /// </summary>
        public int Id { get; set; }
        #endregion

        #region Constructor
        public ListOverride()
        {
            Id = 1;
            ListOverrideCount = 0;
            ListId = 0;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "ID:" + Id + " ListID:" + ListId + " Count:" + ListOverrideCount;
        }
        #endregion
    }
}