using System;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
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

                Email = GetMapiPropertyString(MapiTags.PR_EMAIL_1);

                if (string.IsNullOrEmpty(Email))
                    Email = GetMapiPropertyString(MapiTags.PR_EMAIL_2);

                DisplayName = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);

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
                    {
                        Type = null;
                    }
                    else
                    {
                        switch (recipientType)
                        {
                            case MapiTags.RecipientTo:
                                Type = RecipientType.To;
                                break;

                            case MapiTags.RecipientCC:
                                Type = RecipientType.Cc;
                                break;

                            case MapiTags.RecipientBCC:
                                Type = RecipientType.Bcc;
                                break;

                            case MapiTags.RecipientResource:
                                Type = RecipientType.Resource;
                                break;

                            default:
                                Type = RecipientType.Unknown;
                                break;
                        }
                    }
                }
            }
            #endregion
        }
    }
}