using System;
/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class used to contain To, CC and BCC recipients of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class Recipient : Storage
        {
            #region Public enum RecipientType
            /// <summary>
            /// Recipient types
            /// </summary>
            public enum RecipientType
            {
                /// <summary>
                /// Recipient is unknown
                /// </summary>
                Unknown = 0,
                
                /// <summary>
                /// The recipient is an TO E-mail address
                /// </summary>
                To,

                /// <summary>
                /// The recipient is a CC E-mail address
                /// </summary>
                Cc,

                /// <summary>
                /// The recipient is a BCC E-mail address
                /// </summary>
                Bcc,

                /// <summary>
                /// The recipient is a resource (e.g. a room)
                /// </summary>
                Resource,

                /// <summary>
                ///     The recipient is a room (uses PR_RECIPIENT_TYPE_EXE) needs Exchange 2007 or higher
                /// </summary>
                Room 
            }
            #endregion

            #region Properties
            /// <summary>
            /// Returns the E-mail address, null when not available
            /// </summary>
            public string Email { get; private set; }
            
            /// <summary>
            /// Returns the display name, null when not available
            /// </summary>
            public string DisplayName { get; private set; }

            /// <summary>
            /// Returns the <see cref="RecipientType"/>, null when not available
            /// </summary>
            public RecipientType? Type { get; private set; }

            /// <summary>
            /// Returns the addresstype, null when not available
            /// </summary>
            public string AddressType { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Recipient" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Recipient(Storage message) : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

                string tempEmail;

                AddressType = GetMapiPropertyString(MapiTags.PR_ADDRTYPE);
                if (string.IsNullOrWhiteSpace(AddressType))
                    AddressType = string.Empty;

                switch (AddressType.ToUpperInvariant())
                {
                    case "EX":
                        // In the case of an "Exchange address type" we should try to read the 
                        // PR_SMTP_ADDRESS tag first, otherwhise we have a change to get an E-mail address 
                        // like /O=EXCHANGE/OU=First administrative group/cn=recipients/cn=cmma.van.spelde
                        tempEmail = GetMapiPropertyString(MapiTags.PR_SMTP_ADDRESS);

                        if (string.IsNullOrEmpty(tempEmail))
                            tempEmail = GetMapiPropertyString(MapiTags.PR_EMAIL_ADDRESS);
                        break;

                    default:
                        tempEmail = GetMapiPropertyString(MapiTags.PR_EMAIL_ADDRESS);

                        if (string.IsNullOrEmpty(tempEmail))
                            tempEmail = GetMapiPropertyString(MapiTags.PR_SMTP_ADDRESS);
                        break;
                }

                // If no E-mail address could be found then try to read the very rare PR_ORGEMAILADDR tag
                if (string.IsNullOrEmpty(tempEmail))
                    tempEmail = GetMapiPropertyString(MapiTags.PR_ORGEMAILADDR);

                tempEmail = EmailAddress.RemoveSingleQuotes(tempEmail); 
                var tempDisplayName = EmailAddress.RemoveSingleQuotes(GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME));

                Email = tempEmail;
                DisplayName = tempDisplayName;

                // Sometimes the E-mail address and displayname get swapped so check if they are valid
                if (!EmailAddress.IsEmailAddressValid(tempEmail) && EmailAddress.IsEmailAddressValid(tempDisplayName))
                {
                    // Swap them
                    Email = tempDisplayName;
                    DisplayName = tempEmail;
                }
                else if (EmailAddress.IsEmailAddressValid(tempDisplayName))
                {
                    // If the displayname is an emailAddress them move it
                    Email = tempDisplayName;
                    DisplayName = tempDisplayName;
                }

                if (string.Equals(tempEmail, tempDisplayName, StringComparison.InvariantCultureIgnoreCase))
                    DisplayName = string.Empty;

                // To reliably determine if a recipient is a conference room, use the Messaging API (MAPI) property, PidTagDisplayTypeEx, 
                // of the Recipient object. You can access this property using the PropertyAccessor object in the Outlook object model. 
                // The PidTagDisplayTypeEx property is represented as "http://schemas.microsoft.com/mapi/proptag/0x39050003" in the MAPI 
                // proptag namespace. Note that the PidTagDisplayTypeEx property is not available in versions of Microsoft Exchange Server 
                // earlier than Microsoft Exchange Server 2007; in such earlier versions of Exchange Server, you can use the Recipient.
                // Type property and assume that a recipient having a type other than olResource is not a conference room.
                var displayType = GetMapiPropertyInt32(MapiTags.PR_DISPLAY_TYPE_EX);
                if (displayType != null && displayType == MapiTags.RecipientRoom)
                {
                    Type = RecipientType.Room;
                }
                else
                {
                    var recipientType = GetMapiPropertyInt32(MapiTags.PR_RECIPIENT_TYPE);

                    if (recipientType == null)
                        Type = null;
                    else
                        Type = (RecipientType) recipientType;
                }
            }
            #endregion
        }
    }
}