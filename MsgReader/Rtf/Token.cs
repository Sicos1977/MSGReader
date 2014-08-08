
namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf token type
    /// </summary>
    internal class Token
    {
        #region Properties
        /// <summary>
        /// Type
        /// </summary>
        public RtfTokenType Type { get; set; }

        /// <summary>
        /// Key
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// True when the token has a param
        /// </summary>
        public bool HasParam { get; set; }

        /// <summary>
        /// Param value
        /// </summary>
        public int Param { get; set; }

        // Gives the original hex notation from the Param value when the token key is a [']
        public string Hex { get; set; }

        /// <summary>
        /// Parent token
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public Token Parent { get; set; }

        /// <summary>
        /// True when the token contains text
        /// </summary>
        public bool IsTextToken
        {
            get
            {
                if (Type == RtfTokenType.Text)
                    return true;
                return Type == RtfTokenType.Control && Key == "'" && HasParam;
            }
        }
        #endregion

        #region Constructor
        public Token()
        {
            Type = RtfTokenType.None;
            Parent = null;
            Param = 0;
        }
        #endregion
        
        #region ToString
        public override string ToString()
        {
            if (Type == RtfTokenType.Keyword)
                return Key + Param;

            if (Type == RtfTokenType.GroupStart)
                return "{";

            if (Type == RtfTokenType.GroupEnd)
                return "}";

            if (Type == RtfTokenType.Text)
                return "Text:" + Param;

            return Type + ":" + Key + " " + Param;
        }
        #endregion
    }
}