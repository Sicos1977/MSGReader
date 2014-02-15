namespace DocumentServices.Modules.Readers.MsgReader.Rtf
{
    /// <summary>
    /// Rtf bookmark
    /// </summary>
    public class DomBookmark : DomElement
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