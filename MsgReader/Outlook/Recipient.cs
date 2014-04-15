using System;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {
        /// <summary>
        /// Class used to contain To, CC and BCC recipients of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class Recipient : Storage
        {
            #region Internal enum RecipientType
            /// <summary>
            /// Recipient types
            /// </summary>
            internal enum RecipientType
            {
                /// <summary>
                /// Recipient is a To
                /// </summary>
                To,

                /// <summary>
                /// Recipient is a CC
                /// </summary>
                Cc,

                /// <summary>
                /// Recipient is a BCC
                /// </summary>
                Bcc,

                /// <summary>
                /// Recipient is unknown
                /// </summary>
                Unknown
            }
            #endregion

            #region Properties
            /// <summary>
            /// Gets the display name
            /// </summary>
            public string DisplayName
            {
                get { return GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME); }
            }

            /// <summary>
            /// Gets the recipient email
            /// </summary>
            public string Email
            {
                get
                {
                    var email = GetMapiPropertyString(MapiTags.PR_EMAIL_1);

                    if (string.IsNullOrEmpty(email))
                        email = GetMapiPropertyString(MapiTags.PR_EMAIL_2);

                    return email;
                }
            }

            /// <summary>
            /// Returns true when the <see cref="Storage.Recipient"/> is a room.
            /// This property is only valid when the <see cref="Storage.Message"/> object is an appointment
            /// </summary>
            public bool IsRoom
            {
                get
                {
                    var result = GetMapiPropertyBool(MapiTags.PR_EMAIL_1);
                    if (result != null)
                        return (bool) result;

                    return false;
                }
            }

            /// <summary>
            /// Gets the recipient type
            /// </summary>
            public RecipientType Type
            {
                get
                {
                    var recipientType = GetMapiPropertyInt32(MapiTags.PR_RECIPIENT_TYPE);
                    switch (recipientType)
                    {
                        case MapiTags.MAPI_TO:
                            return RecipientType.To;

                        case MapiTags.MAPI_CC:
                            return RecipientType.Cc;

                        case MapiTags.MAPI_BCC:
                            return RecipientType.Bcc;

                        default:
                            return RecipientType.Unknown;
                    }
                }
            }
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
            }
            #endregion
        }
    }
}