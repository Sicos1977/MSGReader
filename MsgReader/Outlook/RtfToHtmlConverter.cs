﻿//
// RtfToHtmlConverter.cs
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

//using System.Text;

namespace MsgReader.Outlook;

/// <summary>
///     This class is used to convert RTF to HTML
/// </summary>
internal static class RtfToHtmlConverter
{
    #region ConvertRtfToHtml
    /// <summary>
    ///     Convert RTF to HTML
    /// </summary>
    /// <param name="rtf">The rtf string</param>
    /// <returns></returns>
    public static string ConvertRtfToHtml(string rtf)
    {
        if (string.IsNullOrEmpty(rtf))
            return string.Empty;

        var html = RtfPipe.Rtf.ToHtml(rtf.Trim('\0'));
        return html;
    }
    #endregion
}