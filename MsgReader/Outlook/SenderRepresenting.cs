﻿//
// SenderRepresenting.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2024 Kees van Spelde. (www.magic-sessions.com)
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

namespace MsgReader.Outlook;

public partial class Storage
{
    /// <summary>
    ///     Class used to contain the representing sender of a <see cref="Storage.Message" />
    /// </summary>
    public sealed class SenderRepresenting
    {
        #region Fields
        /// <summary>
        ///     The sender in its raw format
        /// </summary>
        private string _raw;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns the E-mail address
        /// </summary>
        public string Email { get; }

        /// <summary>
        ///     Returns the display name
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        ///     Returns the address type, null when not available
        /// </summary>
        public string AddressType { get; }

        /// <summary>
        ///     Returns the <see cref="SenderRepresenting" /> in its raw format
        /// </summary>
        public string Raw
        {
            get
            {
                if (!string.IsNullOrEmpty(_raw))
                    return _raw;

                _raw = string.Empty;

                if (!string.IsNullOrEmpty(DisplayName))
                    _raw = DisplayName.Contains(" ") ? $"\"{DisplayName}\"" : DisplayName;

                if (string.IsNullOrEmpty(Email))
                    return _raw;

                if (!string.IsNullOrEmpty(_raw))
                    _raw += " ";

                _raw += $"<{Email}>";

                return _raw;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     Initializes a new instance of the <see cref="Storage.SenderRepresenting" /> class.
        /// </summary>
        /// <param name="email">The E-mail address of the representing sender</param>
        /// <param name="displayName">The display name of the representing sender</param>
        /// <param name="addressType">The address type</param>
        internal SenderRepresenting(string email, string displayName, string addressType)
        {
            Email = email;
            DisplayName = displayName;
            AddressType = addressType;
        }
        #endregion
    }
}