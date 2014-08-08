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