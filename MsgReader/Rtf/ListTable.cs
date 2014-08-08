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