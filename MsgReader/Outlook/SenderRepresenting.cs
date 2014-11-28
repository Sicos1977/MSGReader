using System;

/*
   Copyright 2013-2014 Kees van Spelde

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
        /// Class used to contain the representing sender of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class SenderRepresenting : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the E-mail address
            /// </summary>
            public string Email { get; private set; }
            
            /// <summary>
            /// Returns the display name
            /// </summary>
            public string DisplayName { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.SenderRepresenting" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal SenderRepresenting(Storage message) : base(message._storage)
            {
                //GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

                var tempEmail = GetMapiPropertyString(MapiTags.PR_SENT_REPRESENTING_EMAIL_ADDRESS);
                tempEmail = EmailAddress.RemoveSingleQuotes(tempEmail); 
                var tempDisplayName = EmailAddress.RemoveSingleQuotes(GetMapiPropertyString(MapiTags.PR_SENT_REPRESENTING_NAME));

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
            }
            #endregion
        }
    }
}