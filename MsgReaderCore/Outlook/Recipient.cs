//
// Recipient.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    #region Enum RecipientType
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

    public partial class Storage
    {
        /// <summary>
        /// Class used to contain To, CC and BCC recipients of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class Recipient : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the E-mail address, null when not available
            /// </summary>
            public string Email { get; }
            
            /// <summary>
            /// Returns the display name, null when not available
            /// </summary>
            public string DisplayName { get; }

            /// <summary>
            /// Returns the <see cref="RecipientType"/>, null when not available
            /// </summary>
            public RecipientType? Type { get; }

            /// <summary>
            /// Returns the addresstype, null when not available
            /// </summary>
            public string AddressType { get; }
            #endregion

            #region Constructor
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Recipient" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Recipient(Storage message) : base(message._rootStorage)
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

                // If no E-mail address could be found then try to read the very rare PR_ORGEMAILADDR tag , override original if it does not contain an @
                if (string.IsNullOrEmpty(tempEmail) || !tempEmail.Contains("@"))
                {
                    var testEmail = GetMapiPropertyString(MapiTags.PR_ORGEMAILADDR);
                    if (!string.IsNullOrEmpty(testEmail) && testEmail.Contains("@"))
                    {
                        tempEmail = testEmail;
                    }
                }

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