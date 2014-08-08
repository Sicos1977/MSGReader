namespace MsgReader.Rtf
{
    /// <summary>
    /// Document field element
    /// </summary>
    internal class DomField : DomElement
    {
        #region Constructor
        public DomField()
        {
            Method = RtfDomFieldMethod.None;
        }
        #endregion

        #region Properties
        // ReSharper disable UnusedMember.Global
        // ReSharper disable UnusedAutoPropertyAccessor.Global
        /// <summary>
        /// Method
        /// </summary>
        public RtfDomFieldMethod Method { get; set; }

        /// <summary>
        /// Instructions
        /// </summary>
        public string Instructions
        // ReSharper restore UnusedMember.Global
        {
            get
            {
                foreach (DomElement element in Elements)
                {
                    var container = element as ElementContainer;
                    if (container != null)
                    {
                        if (container.Name == Consts.Fldinst)
                            return container.InnerText;
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Result
        /// </summary>
        public ElementContainer Result
        {
            get
            {
                foreach (DomElement element in Elements)
                {
                    var container = element as ElementContainer;
                    if (container != null)
                    {
                        if (container.Name == Consts.Fldrslt)
                            return container;
                    }
                }
                return null;
            }
        }

        public string ResultString
        {
            get
            {
                return Result != null ? Result.InnerText : null;
            }
        }
        // ReSharper restore UnusedAutoPropertyAccessor.Global
        // ReSharper disable once MemberCanBePrivate.Global
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Field"; // +strInstructions + " Result:" + this.ResultString;
        }
        #endregion
    }
}