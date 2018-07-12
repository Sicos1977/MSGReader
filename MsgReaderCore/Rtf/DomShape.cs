namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf shape element
    /// </summary>
    internal class DomShape : DomElement
    {
        #region Properties
        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable UnusedAutoPropertyAccessor.Global
        /// <summary>
        /// Left position
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Top position
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Z index
        /// </summary>
        public int ZIndex { get; set; }

        /// <summary>
        /// Shape id
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Ext attribute
        /// </summary>
        public StringAttributeCollection ExtAttrbutes { get; set; }
        // ReSharper restore UnusedAutoPropertyAccessor.Global
        // ReSharper restore MemberCanBePrivate.Global
        #endregion

        #region Constructor
        public DomShape()
        {
            ExtAttrbutes = new StringAttributeCollection();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Shape:Left:" + Left + " Top:" + Top + " Width:" + Width + " Height:" + Height;
        }
        #endregion
    }
}