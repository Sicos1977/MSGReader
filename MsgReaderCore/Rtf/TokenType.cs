//
// Enums.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2023 Magic-Sessions. (www.magic-sessions.com)
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

namespace MsgReader.Rtf;

#region Enum RtfTokenType
/// <summary>
///     Rtf token type
/// </summary>
internal enum TokenType
{
    /// <summary>
    ///     Nothing
    /// </summary>
    None,
    
    /// <summary>
    ///     A RTF keyword for example /lang /f , etc...
    /// </summary>
    Keyword,
    
    /// <summary>
    ///     Extension keyword (for HTML or Text extraction)
    /// </summary>
    ExtensionKeyword,
    
    /// <summary>
    ///     A control char e.g. ' or u
    /// </summary>
    Control,

    /// <summary>
    ///     Just plain text
    /// </summary>
    Text,

    /// <summary>
    ///     The end of the file
    /// </summary>
    Eof,

    /// <summary>
    ///     A group start {
    /// </summary>
    GroupStart,

    /// <summary>
    ///     A group end }
    /// </summary>
    GroupEnd
}
#endregion