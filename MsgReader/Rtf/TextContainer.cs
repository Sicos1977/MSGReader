//
// TextContainer.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Kees van Spelde. (www.magic-sessions.com)
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

using System.Text;

namespace MsgReader.Rtf;

/// <summary>
///     Rtf plain text container
/// </summary>
internal class TextContainer
{
    #region Fields
    private readonly ByteBuffer _byteBuffer = new();
    private readonly StringBuilder _stringBuilder = new();
    private readonly Encoding _encoding;
    #endregion

    #region Properties
    /// <summary>
    ///     text value
    /// </summary>
    public string Text => _stringBuilder.ToString();
    #endregion

    #region Constructor
    /// <summary>
    ///     Initialize instance
    /// </summary>
    /// <param name="encoding"></param>
    public TextContainer(Encoding encoding)
    {
        _encoding = encoding;
    }
    #endregion

    #region Append
    /// <summary>
    ///     Append text content
    /// </summary>
    /// <param name="text"></param>
    public void Append(string text)
    {
        if (!string.IsNullOrEmpty(text))
            _stringBuilder.Append(text);
    }
    #endregion

    #region Accept
    /// <summary>
    ///     Accept rtf token
    /// </summary>
    /// <param name="token">RTF token</param>
    /// <param name="reader"></param>
    public void Accept(Token token, Reader reader)
    {
        if (token == null) return;

        if (_byteBuffer.Count > 0 && reader.TokenType != TokenType.EncodedChar)
        {
            _stringBuilder.Append(_byteBuffer.GetString(_encoding));
            _byteBuffer.Clear();
        }

        switch (token.Type)
        {
            case TokenType.Text:
                _stringBuilder.Append(token.Key);
                break;

            case TokenType.EncodedChar:
                if (token.HasParam)
                {
                    switch (token.Key)
                    {
                        case Consts.U:
                            _stringBuilder.Append((char)token.Param);
                            reader.CurrentLayerInfo.UcValueCount = reader.CurrentLayerInfo.UcValue;
                            return;

                        case Consts.Apostrophe:
                            _byteBuffer.Add((byte)token.Param);
                            return;
                    }
                }

                break;
        }

        if (token.Key == Consts.Tab)
        {
            _stringBuilder.Append("\t");
            return;
        }

        if (token.Key == Consts.Emdash)
        {
            _stringBuilder.Append('-');
            return;
        }

        if (token.Key == "")
            _stringBuilder.Append('-');
    }
    #endregion
}