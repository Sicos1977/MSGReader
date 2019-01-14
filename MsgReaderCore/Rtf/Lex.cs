//
// Lex.cs
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
using System.Globalization;
using System.IO;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    ///     Rtf lex
    /// </summary>
    internal class Lex
    {
        #region Fields
        private const int Eof = -1;
        private readonly TextReader _reader;
        #endregion

        #region Constructor
        /// <summary>
        ///     Initialize instance
        /// </summary>
        /// <param name="reader">reader</param>
        public Lex(TextReader reader)
        {
            _reader = reader;
        }
        #endregion

        #region PeekTokenType
        /// <summary>
        ///     Peek to see what kind of token we have
        /// </summary>
        /// <returns>TokenType</returns>
        public RtfTokenType PeekTokenType()
        {
            var c = _reader.Peek();

            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
            {
                _reader.Read();
                c = _reader.Peek();
            }

            if (c == Eof)
                return RtfTokenType.Eof;

            switch (c)
            {
                case '{':
                    return RtfTokenType.GroupStart;

                case '}':
                    return RtfTokenType.GroupEnd;

                case '\\':
                    return RtfTokenType.Control;

                default:
                    return RtfTokenType.Text;
            }
        }
        #endregion

        #region NextToken
        /// <summary>
        ///     Read next token
        /// </summary>
        /// <returns>token</returns>
        public Token NextToken()
        {
            var token = new Token();

            var c = _reader.Read();

            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
                c = _reader.Read();

            if (c != Eof)
                switch (c)
                {
                    case '{':
                        token.Type = RtfTokenType.GroupStart;
                        break;

                    case '}':
                        token.Type = RtfTokenType.GroupEnd;
                        break;

                    case '\\':
                        ParseKeyword(token);
                        break;

                    default:
                        token.Type = RtfTokenType.Text;
                        ParseText(c, token);
                        break;
                }
            else
                token.Type = RtfTokenType.Eof;

            return token;
        }
        #endregion

        #region ParseKeyword
        /// <summary>
        ///     Parse keyword from token
        /// </summary>
        /// <param name="token"></param>
        private void ParseKeyword(Token token)
        {
            var ext = false;
            var c = _reader.Peek();

            if (!char.IsLetter((char) c))
            {
                _reader.Read();
                if (c == '*')
                {
                    // Expand keyword
                    token.Type = RtfTokenType.Keyword;
                    _reader.Read();
                    ext = true;
                }
                else
                {
                    if (c == '\\' || c == '{' || c == '}')
                    {
                        // Special character
                        token.Type = RtfTokenType.Text;
                        token.Key = ((char) c).ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        token.Type = RtfTokenType.Control;
                        token.Key = ((char) c).ToString(CultureInfo.InvariantCulture);

                        if (token.Key == "\'")
                        {
                            // Read 2 hex characters
                            var text = new StringBuilder();
                            text.Append((char) _reader.Read());
                            text.Append((char) _reader.Read());
                            token.HasParam = true;
                            token.Hex = text.ToString().ToLower();

                            token.Param = Convert.ToInt32(text.ToString().ToLower(), 16);
                        }
                    }

                    return;
                }
            }

            // Read keyword
            var keyword = new StringBuilder();
            c = _reader.Peek();

            while (char.IsLetter((char) c))
            {
                _reader.Read();
                keyword.Append((char) c);
                c = _reader.Peek();
            }

            token.Type = ext ? RtfTokenType.ExtKeyword : RtfTokenType.Keyword;
            token.Key = keyword.ToString();

            // Read an integer
            if (char.IsDigit((char) c) || c == '-')
            {
                token.HasParam = true;
                var negative = false;

                if (c == '-')
                {
                    negative = true;
                    _reader.Read();
                }

                c = _reader.Peek();

                var text = new StringBuilder();

                while (char.IsDigit((char) c))
                {
                    _reader.Read();
                    text.Append((char) c);
                    c = _reader.Peek();
                }

                var param = Convert.ToInt32(text.ToString());
                if (negative)
                    param = -param;

                token.Param = param;
            }

            if (c == ' ')
                _reader.Read();
        }
        #endregion

        #region ParseText
        /// <summary>
        ///     Parse text after char
        /// </summary>
        /// <param name="c"></param>
        /// <param name="token"></param>
        private void ParseText(int c, Token token)
        {
            var stringBuilder = new StringBuilder(((char) c).ToString(CultureInfo.InvariantCulture));

            c = ClearWhiteSpace();

            while (c != '\\' && c != '}' && c != '{' && c != Eof)
            {
                _reader.Read();
                stringBuilder.Append((char) c);
                c = ClearWhiteSpace();
            }

            token.Key = stringBuilder.ToString();
        }

        /// <summary>
        ///     Read chars until another non white space char is found
        /// </summary>
        /// <returns></returns>
        private int ClearWhiteSpace()
        {
            var c = _reader.Peek();
            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
            {
                _reader.Read();
                c = _reader.Peek();
            }

            return c;
        }
        #endregion
    }
}