namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf bookmark
    /// </summary>
    internal class DomBookmark : DomElement
    {
        #region Properties
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "BookMark:" + Name;
        }
        #endregion
    }
}