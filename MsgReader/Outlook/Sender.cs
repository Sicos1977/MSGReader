using System;
using System.Text;
using System.Text.RegularExpressions;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {
        /// <summary>
        /// Class used to contain the Sender of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class Sender : Storage
        {
            #region Properties
            /// <summary>
            /// Gets the display value of the contact that sent the email.
            /// </summary>
            public string DisplayName
            {
                get { return GetMapiPropertyString(MapiTags.PR_SENDER_NAME); }
            }

            /// <summary>
            /// Gets the sender email
            /// </summary>
            public string Email
            {
                get
                {
                    var eMail = GetMapiPropertyString(MapiTags.PR_SENDER_EMAIL_ADDRESS);

                    if (string.IsNullOrEmpty(eMail) || eMail.IndexOf('@') < 0)
                        eMail = GetMapiPropertyString(MapiTags.PR_SENDER_EMAIL_ADDRESS_2);

                    if (string.IsNullOrEmpty(eMail) || eMail.IndexOf("@", StringComparison.Ordinal) < 0)
                    {
                        // Get address from email header
                        var header = GetStreamAsString(MapiTags.HeaderStreamName, Encoding.Unicode);
                        var m = Regex.Match(header, "From:.*<(?<email>.*?)>");
                        eMail = m.Groups["email"].ToString();
                    }

                    return eMail;
                }
            }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Sender" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Sender(Storage message)
                : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;
            }
            #endregion
        }
    }
}