//
// ReceivedBy.cs
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
            public string AddressType { get; }
            
            /// <summary>
            /// Returns the E-mail address, null when not available
            /// </summary>
            public string Email { get; }

            /// <summary>
            /// Returns the display name, null when not available
            /// </summary>
            public string DisplayName { get; }
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
