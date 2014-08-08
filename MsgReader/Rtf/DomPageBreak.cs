namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf dom page break
    /// </summary>
    internal class DomPageBreak : DomElement
    {
        #region Properties
        public override string InnerText
        {
            get { return string.Empty; }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize this instance
        /// </summary>
        public DomPageBreak()
        {
            Locked = true;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "page";
        }
        #endregion
    }
}