using System;
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// This class contains information about the person that has received this message
        /// </summary>
        public sealed class ReceivedBy : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the address type (e.g. SMTP)
            /// </summary>
            public string AddressType { get; private set; }
            
            /// <summary>
            /// Returns the E-mail address, null when not available
            /// </summary>
            public string Email { get; private set; }

            /// <summary>
            /// Returns the display name, null when not available
            /// </summary>
            public string DisplayName { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.ReceivedBy" /> class.
            /// </summary>
            internal ReceivedBy(string addressType, string email, string displayName)
            {
                AddressType = addressType;
                Email = email;
                DisplayName = displayName;
                
                var tempEmail = EmailAddress.RemoveSingleQuotes(Email);
                var tempDisplayName = EmailAddress.RemoveSingleQuotes(DisplayName);

                Email = tempEmail;
                DisplayName = tempDisplayName;

                // Sometimes the E-mail address and displayname get swapped so check if they are valid
                if (!EmailAddress.IsEmailAddressValid(tempEmail) && EmailAddress.IsEmailAddressValid(tempDisplayName))
                {
                    // Swap then
                    Email = tempDisplayName;
                    DisplayName = tempEmail;
                }
                else if (EmailAddress.IsEmailAddressValid(tempDisplayName))
                {
                    // If the displayname is an emailAddress then move it
                    Email = tempDisplayName;
                    DisplayName = tempDisplayName;
                }

                if (string.Equals(tempEmail, tempDisplayName, StringComparison.InvariantCultureIgnoreCase))
                    DisplayName = string.Empty;
            }
            #endregion
        }
    }
}
