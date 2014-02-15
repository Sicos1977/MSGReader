namespace DocumentServices.Modules.Readers.MsgReader.Rtf
{
    /// <summary>
    /// RTF element container
    /// </summary>
    public class ElementContainer : DomElement
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
            return "Container : " + Name;
        }
        #endregion
    }
}