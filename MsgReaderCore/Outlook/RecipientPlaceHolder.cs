//
// RecipientPlaceHolder.cs
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

namespace MsgReader.Outlook
{
    /// <summary>
    ///     Used as a placeholder for the recipients from the MSG file itself or from the "internet"
    ///     headers when this message is send outside an Exchange system
    /// </summary>
    public sealed class RecipientPlaceHolder
    {
        #region Properties
        /// <summary>
        ///     The E-mail address
        /// </summary>
        public string Email { get; }

        /// <summary>
        ///     The display name
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        /// Returns the addresstype, null when not available
        /// </summary>
        public string AddressType { get; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all it's properties
        /// </summary>
        /// <param name="email">The E-mail address</param>
        /// <param name="displayName">The display name</param>
        /// <param name="addressType">The address type</param>
        internal RecipientPlaceHolder(string email, string displayName, string addressType)
        {
            Email = email;
            DisplayName = displayName;
            AddressType = addressType;
        }
        #endregion
    }
}