//
// ListTable.cs
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
    /// Rtf list table
    /// </summary>
    internal class ListTable : List<RtfList>
    {
        #region GetById
        public RtfList GetById(int id)
        {
            foreach (var rtfList in this)
            {
                if (rtfList.ListId == id)
                    return rtfList;
            }
            return null;
        }
        #endregion
    }

    /// <summary>
    /// Rtf list table item
    /// </summary>
    internal class RtfList
    {
        #region Properties
        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable UnusedAutoPropertyAccessor.Global
        public int ListId { get; set; }

        public int ListTemplateId { get; set; }

        public bool ListSimple { get; set; }

        public bool ListHybrid { get; set; }

        public string ListName { get; set; }

        public string ListStyleName { get; set; }

        public int LevelStartAt { get; set; }

        public RtfLevelNumberType LevelNfc { get; set; }

        public int LevelJc { get; set; }

        public int LevelFollow { get; set; }

        public string FontName { get; set; }

        public string LevelText { get; set; }
        // ReSharper restore UnusedAutoPropertyAccessor.Global
        // ReSharper restore MemberCanBePrivate.Global
        #endregion

        #region Constructor
        public RtfList()
        {
            ListId = 0;
            LevelText = null;
            FontName = null;
            LevelFollow = 0;
            LevelJc = 0;
            LevelNfc = RtfLevelNumberType.None;
            LevelStartAt = 1;
            ListName = null;
            ListSimple = false;
            ListTemplateId = 0;
            ListHybrid = false;
            ListStyleName = null;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            if (LevelNfc == RtfLevelNumberType.Bullet)
            {
                var text = "ID:" + ListId + "   Bullet:";
                if (string.IsNullOrEmpty(LevelText) == false)
                {
                    text = text + "(" + Convert.ToString((short) LevelText[0]) + ")";
                }
                return text;
            }

            return "ID:" + ListId + " " + LevelNfc.ToString() + " Start:" + LevelStartAt;
        }
        #endregion
    }
}