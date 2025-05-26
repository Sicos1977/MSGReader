//
// EmlDateParser.cs
//
// Author: siemvanoers
// 
// Copyright (c) 2013-2025 Siem
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
using System.Text.RegularExpressions;

namespace MsgReader.Helpers

/// <summary>
///     Helper class for parsing the Date header from an EML file,
///     preserving the time zone offset as specified in the header.
/// </summary>
{
    public static class EmlDateParser
    {
        #region ParseDateHeader
        /// <summary>
        ///     Extracts and parses the Date header from the given EML content,
        ///     preserving the time zone offset.
        /// </summary>
        /// <param name="emlContent">The raw EML content as a string.</param>
        /// <returns>
        ///     A <see cref="DateTimeOffset"/> representing the parsed date and time zone,
        ///     or <c>null</c> if the Date header is not found or cannot be parsed.
        /// </returns>
        public static DateTimeOffset? ParseDateHeader(string emlContent)
        {
            // Extract the Date header (case-insensitive)
            var match = Regex.Match(emlContent, @"^Date:\s*(.+)$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            if (!match.Success)
                return null;

            var dateHeader = match.Groups[1].Value.Trim();

            // Parse using DateTimeOffset to preserve the offset
            if (DateTimeOffset.TryParse(dateHeader, out var parsedDate))
                return parsedDate;

            return null;
        }
        #endregion
    }
}
