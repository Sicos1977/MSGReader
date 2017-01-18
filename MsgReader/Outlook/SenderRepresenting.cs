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

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class used to contain the representing sender of a <see cref="Storage.Message"/>
        /// </summary>
        public sealed class SenderRepresenting
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

            /// <summary>
            /// Returns the addresstype, null when not available
            /// </summary>
            public string AddressType { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.SenderRepresenting" /> class.
            /// </summary>
            /// <param name="email">The E-mail address of the representing sender</param>
            /// <param name="displayName">The displayname of the representing sender</param>
            /// <param name="addresType">The address type</param>
            internal SenderRepresenting(string email, string displayName, string addresType)
            {
                Email = email;
                DisplayName = displayName;
                AddressType = addresType;
            }
            #endregion
        }
    }
}