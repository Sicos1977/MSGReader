namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF element container
    /// </summary>
    internal class ElementContainer : DomElement
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