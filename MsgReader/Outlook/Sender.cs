using System;
using System.Text;
using System.Text.RegularExpressions;

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

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class used to contain the Sender of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class Sender : Storage
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
            /// Initializes a new instance of the <see cref="Storage.Sender" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Sender(Storage message) : base(message._storage)
            {
                //GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

                Email = GetMapiPropertyString(MapiTags.PR_SENDER_EMAIL_ADDRESS);

                if (string.IsNullOrEmpty(Email) || Email.IndexOf('@') < 0)
                    Email = GetMapiPropertyString(MapiTags.PR_SENDER_EMAIL_ADDRESS_2);

                if (string.IsNullOrEmpty(Email) || Email.IndexOf("@", StringComparison.Ordinal) < 0)
                {
                    var addressType = GetMapiPropertyString(MapiTags.PR_SENDER_ADDRTYPE);
                    if (addressType == null || addressType == "EX")
                        Email = null;
                    else
                    {
                        // Get address from email headers. The headers are not present when the addressType = "EX"
                        var header = GetStreamAsString(MapiTags.HeaderStreamName, Encoding.Unicode);
                        if (header == null) return;
                        var matches = Regex.Match(header, "From:.*<(?<email>.*?)>");
                        Email = matches.Groups["email"].ToString();
                    }
                }

                DisplayName = GetMapiPropertyString(MapiTags.PR_SENDER_NAME);
            }
            #endregion
        }
    }
}