namespace MsgReader.Outlook
{
    /// <summary>
    ///     Used as a placeholder for the recipients from the MSG file itself or from the "internet"
    ///     headers when this message is send outside an Exchange system
    /// </summary>
    public sealed class RecipientPlaceHolder
    {
        #region Properties
        /// <summary>
        ///     The E-mail address
        /// </summary>
        public string Email { get; private set; }

        /// <summary>
        ///     The display name
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        /// Returns the addresstype, null when not available
        /// </summary>
        public string AddressType { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all it's properties
        /// </summary>
        /// <param name="email">The E-mail address</param>
        /// <param name="displayName">The display name</param>
        /// <param name="addressType">The address type</param>
        internal RecipientPlaceHolder(string email, string displayName, string addressType)
        {
            Email = email;
            DisplayName = displayName;
            AddressType = addressType;
        }
        #endregion
    }
}