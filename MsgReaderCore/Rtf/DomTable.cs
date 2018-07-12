namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf table
    /// </summary>
    internal class DomTable : DomElement
    {
        #region Properties
        /// <summary>
        /// column list
        /// </summary>
        public DomElementList Columns { get; set; }
        #endregion

        #region Constructor
        public DomTable()
        {
            Columns = new DomElementList();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Table(Rows:" + Elements.Count + " Columns:" + Columns.Count + ")";
        }
        #endregion
    }
}