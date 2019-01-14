//
// Token.cs
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

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf token type
    /// </summary>
    internal class Token
    {
        #region Properties
        /// <summary>
        /// Type
        /// </summary>
        public RtfTokenType Type { get; set; }

        /// <summary>
        /// Key
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// True when the token has a param
        /// </summary>
        public bool HasParam { get; set; }

        /// <summary>
        /// Param value
        /// </summary>
        public int Param { get; set; }

        // Gives the original hex notation from the Param value when the token key is a [']
        public string Hex { get; set; }

        /// <summary>
        /// Parent token
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public Token Parent { get; set; }

        /// <summary>
        /// True when the token contains text
        /// </summary>
        public bool IsTextToken
        {
            get
            {
                if (Type == RtfTokenType.Text)
                    return true;
                return Type == RtfTokenType.Control && Key == "'" && HasParam;
            }
        }
        #endregion

        #region Constructor
        public Token()
        {
            Type = RtfTokenType.None;
            Parent = null;
            Param = 0;
        }
        #endregion
        
        #region ToString
        public override string ToString()
        {
            if (Type == RtfTokenType.Keyword)
                return Key + Param;

            if (Type == RtfTokenType.GroupStart)
                return "{";

            if (Type == RtfTokenType.GroupEnd)
                return "}";

            if (Type == RtfTokenType.Text)
                return "Text:" + Param;

            return Type + ":" + Key + " " + Param;
        }
        #endregion
    }
}