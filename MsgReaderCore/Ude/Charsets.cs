/* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1/GPL 2.0/LGPL 2.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is Mozilla Universal charset detector code.
 *
 * The Initial Developer of the Original Code is
 * Netscape Communications Corporation.
 * Portions created by the Initial Developer are Copyright (C) 2001
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *          Kohei TAKETA <k-tak@void.in> (Java port)
 *          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
 *
 * Refactoring to the code done by Kees van Spelde so that it works in this project
 * Copyright (c) 2023 Magic-Sessions. (www.magic-sessions.com)
 *
 * Alternatively, the contents of this file may be used under the terms of
 * either the GNU General Public License Version 2 or later (the "GPL"), or
 * the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
 * in which case the provisions of the GPL or the LGPL are applicable instead
 * of those above. If you wish to allow use of your version of this file only
 * under the terms of either the GPL or the LGPL, and not to allow others to
 * use your version of this file under the terms of the MPL, indicate your
 * decision by deleting the provisions above and replace them with the notice
 * and other provisions required by the GPL or the LGPL. If you do not delete
 * the provisions above, a recipient may use your version of this file under
 * the terms of any one of the MPL, the GPL or the LGPL.
 *
 * ***** END LICENSE BLOCK ***** */

// ReSharper disable UnusedMember.Global
namespace MsgReader.Ude;

internal static class Charsets
{
    #region Consts
    public const string Ascii = "ASCII";

    public const string Utf8 = "UTF-8";

    public const string Utf16Le = "UTF-16LE";

    public const string Utf16Be = "UTF-16BE";

    public const string Utf32Be = "UTF-32BE";

    public const string Utf32Le = "UTF-32LE";

    /// <summary>
    ///     Unusual BOM (3412 order)
    /// </summary>
    public const string Ucs43412 = "X-ISO-10646-UCS-4-3412";

    /// <summary>
    ///     Unusual BOM (2413 order)
    /// </summary>
    public const string Ucs42413 = "X-ISO-10646-UCS-4-2413";

    /// <summary>
    ///     Cyrillic (based on bulgarian and russian data)
    /// </summary>
    public const string Win1251 = "windows-1251";

    /// <summary>
    ///     Latin-1, almost identical to ISO-8859-1
    /// </summary>
    public const string Win1252 = "windows-1252";

    /// <summary>
    ///     Greek
    /// </summary>
    public const string Win1253 = "windows-1253";

    /// <summary>
    ///     Logical hebrew (includes ISO-8859-8-I and most of x-mac-hebrew)
    /// </summary>
    public const string Win1255 = "windows-1255";

    /// <summary>
    ///     Traditional chinese
    /// </summary>
    public const string Big5 = "Big5";

    public const string Euckr = "EUC-KR";

    public const string Eucjp = "EUC-JP";

    public const string Euctw = "EUC-TW";

    /// <summary>
    ///     Note: gb2312 is a subset of gb18030
    /// </summary>
    public const string Gb18030 = "gb18030";

    public const string Iso2022Jp = "ISO-2022-JP";

    public const string Iso2022Cn = "ISO-2022-CN";

    public const string Iso2022Kr = "ISO-2022-KR";

    /// <summary>
    ///     Simplified chinese
    /// </summary>
    public const string HzGb2312 = "HZ-GB-2312";

    public const string ShiftJis = "Shift-JIS";

    public const string MacCyrillic = "x-mac-cyrillic";

    public const string Koi8R = "KOI8-R";

    public const string Ibm855 = "IBM855";

    public const string Ibm866 = "IBM866";

    /// <summary>
    ///     East-Europe. Disabled because too similar to windows-1252
    ///     (latin-1). Should use tri-grams models to discriminate between
    ///     these two charsets.
    /// </summary>
    public const string Iso88592 = "ISO-8859-2";

    /// <summary>
    ///     Cyrillic
    /// </summary>
    public const string Iso88595 = "ISO-8859-5";

    /// <summary>
    ///     Greek
    /// </summary>
    public const string Iso88597 = "ISO-8859-7";

    /// <summary>
    ///     Visual Hebrew
    /// </summary>
    public const string Iso88598 = "ISO-8859-8";

    /// <summary>
    ///     Thai. This recognizer is not enabled yet.
    /// </summary>
    public const string Tis620 = "TIS620";
    #endregion
}