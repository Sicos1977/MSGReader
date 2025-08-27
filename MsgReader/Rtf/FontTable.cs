//
// FontTable.cs
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

using System.Collections.Generic;
using System.Text;
// ReSharper disable StringLiteralTypo

namespace MsgReader.Rtf;

/// <summary>
///     Rtf font information
/// </summary>
internal class Font
{
    #region Consts
    private const int AnsiCodePage = 1252;
    private const int SymbolFakeCodePage = 42; // a windows legacy hack ...
    #endregion

    #region Fields
    private int _charset;

    private static readonly Dictionary<int, string> CodePages = new()
    {
        { 37, "IBM037" },
        { 437, "IBM437" },
        { 500, "IBM500" },
        { 708, "ASMO-708" },
        { 720, "DOS-720" },
        { 737, "ibm737" },
        { 775, "ibm775" },
        { 850, "ibm850" },
        { 852, "ibm852" },
        { 855, "IBM855" },
        { 857, "ibm857" },
        { 858, "IBM00858" },
        { 860, "IBM860" },
        { 861, "ibm861" },
        { 862, "DOS-862" },
        { 863, "IBM863" },
        { 864, "IBM864" },
        { 865, "IBM865" },
        { 866, "cp866" },
        { 869, "ibm869" },
        { 870, "IBM870" },
        { 874, "windows-874" },
        { 875, "cp875" },
        { 932, "shift_jis" },
        { 936, "gb2312" },
        { 949, "ks_c_5601-1987" },
        { 950, "big5" },
        { 1026, "IBM1026" },
        { 1047, "IBM01047" },
        { 1140, "IBM01140" },
        { 1141, "IBM01141" },
        { 1142, "IBM01142" },
        { 1143, "IBM01143" },
        { 1144, "IBM01144" },
        { 1145, "IBM01145" },
        { 1146, "IBM01146" },
        { 1147, "IBM01147" },
        { 1148, "IBM01148" },
        { 1149, "IBM01149" },
        { 1200, "utf-16" },
        { 1201, "utf-16BE" },
        { 1250, "windows-1250" },
        { 1251, "windows-1251" },
        { 1252, "windows-1252" },
        { 1253, "windows-1253" },
        { 1254, "windows-1254" },
        { 1255, "windows-1255" },
        { 1256, "windows-1256" },
        { 1257, "windows-1257" },
        { 1258, "windows-1258" },
        { 1361, "Johab" },
        { 10000, "macintosh" },
        { 10001, "x-mac-japanese" },
        { 10002, "x-mac-chinesetrad" },
        { 10003, "x-mac-korean" },
        { 10004, "x-mac-arabic" },
        { 10005, "x-mac-hebrew" },
        { 10006, "x-mac-greek" },
        { 10007, "x-mac-cyrillic" },
        { 10008, "x-mac-chinesesimp" },
        { 10010, "x-mac-romanian" },
        { 10017, "x-mac-ukrainian" },
        { 10021, "x-mac-thai" },
        { 10029, "x-mac-ce" },
        { 10079, "x-mac-icelandic" },
        { 10081, "x-mac-turkish" },
        { 10082, "x-mac-croatian" },
        { 12000, "utf-32" },
        { 12001, "utf-32BE" },
        { 20000, "x-Chinese-CNS" },
        { 20001, "x-cp20001" },
        { 20002, "x-Chinese-Eten" },
        { 20003, "x-cp20003" },
        { 20004, "x-cp20004" },
        { 20005, "x-cp20005" },
        { 20105, "x-IA5" },
        { 20106, "x-IA5-German" },
        { 20107, "x-IA5-Swedish" },
        { 20108, "x-IA5-Norwegian" },
        { 20127, "us-ascii" },
        { 20261, "x-cp20261" },
        { 20269, "x-cp20269" },
        { 20273, "IBM273" },
        { 20277, "IBM277" },
        { 20278, "IBM278" },
        { 20280, "IBM280" },
        { 20284, "IBM284" },
        { 20285, "IBM285" },
        { 20290, "IBM290" },
        { 20297, "IBM297" },
        { 20420, "IBM420" },
        { 20423, "IBM423" },
        { 20424, "IBM424" },
        { 20833, "x-EBCDIC-KoreanExtended" },
        { 20838, "IBM-Thai" },
        { 20866, "koi8-r" },
        { 20871, "IBM871" },
        { 20880, "IBM880" },
        { 20905, "IBM905" },
        { 20924, "IBM00924" },
        { 20932, "EUC-JP" },
        { 20936, "x-cp20936" },
        { 20949, "x-cp20949" },
        { 21025, "cp1025" },
        { 21866, "koi8-u" },
        { 28591, "iso-8859-1" },
        { 28592, "iso-8859-2" },
        { 28593, "iso-8859-3" },
        { 28594, "iso-8859-4" },
        { 28595, "iso-8859-5" },
        { 28596, "iso-8859-6" },
        { 28597, "iso-8859-7" },
        { 28598, "iso-8859-8" },
        { 28599, "iso-8859-9" },
        { 28603, "iso-8859-13" },
        { 28605, "iso-8859-15" },
        { 29001, "x-Europa" },
        { 38598, "iso-8859-8-i" },
        { 50220, "iso-2022-jp" },
        { 50221, "csISO2022JP" },
        { 50222, "iso-2022-jp" },
        { 50225, "iso-2022-kr" },
        { 50227, "x-cp50227" },
        { 51932, "euc-jp" },
        { 51936, "EUC-CN" },
        { 51949, "euc-kr" },
        { 52936, "hz-gb-2312" },
        { 54936, "GB18030" },
        { 57002, "x-iscii-de" },
        { 57003, "x-iscii-be" },
        { 57004, "x-iscii-ta" },
        { 57005, "x-iscii-te" },
        { 57006, "x-iscii-as" },
        { 57007, "x-iscii-or" },
        { 57008, "x-iscii-ka" },
        { 57009, "x-iscii-ma" },
        { 57010, "x-iscii-gu" },
        { 57011, "x-iscii-pa" },
        { 65000, "utf-7" },
        { 65001, "utf-8" }
    };
    #endregion

    #region Properties
    /// <summary>
    ///     The default encoding
    /// </summary>
    public static Encoding DefaultEncoding { get; } = Encoding.GetEncoding("Windows-1252");

    /// <summary>
    ///     Font index
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    ///     font name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    ///     Charset
    /// </summary>
    public int Charset
    {
        get => _charset;
        set
        {
            _charset = value;
            Encoding = EncodingFromCharset(_charset);
        }
    }

    /// <summary>
    ///     Encoding
    /// </summary>
    public Encoding Encoding { get; private set; }
    #endregion

    #region Clone
    public Font Clone()
    {
        return new Font
        {
            _charset = _charset,
            Index = Index,
            Encoding = Encoding,
            Name = Name
        };
    }
    #endregion

    #region ToString
    public override string ToString()
    {
        return $"{Index}:{Name} Charset:{_charset}";
    }
    #endregion

    #region EncodingFromCharset
    /// <summary>
    ///     Returns an encoding for a given <paramref name="charset"/>
    /// </summary>
    /// <param name="charset"></param>
    /// <returns></returns>
    public static Encoding EncodingFromCharset(int charset)
    {
        return EncodingFromCodePage(GetCodePage(charset));
    }
    #endregion

    #region EncodingFromCodePage
    public static Encoding EncodingFromCodePage(int codePage)
    {
        switch (codePage)
        {
            case 0:
            case AnsiCodePage:
            case SymbolFakeCodePage:
                return DefaultEncoding;
            default:
                return Encoding.GetEncoding(CodePages[codePage]);
        }
    }
    #endregion

    #region GetCodePage
    /// <summary>
    ///     Returns the codepage for the given <paramref name="charset"/>
    /// </summary>
    /// <param name="charset"></param>
    /// <returns></returns>
    public static int GetCodePage(int charset)
    {
        switch (charset)
        {
            case 0:
                return 1252; // ANSI
            case 1:
                return 0; // Default
            case 2:
                return 42; // Symbol
            case 77:
                return 10000; // Mac Roman
            case 78:
                return 10001; // Mac Shift Jis
            case 79:
                return 10003; // Mac Hangul
            case 80:
                return 10008; // Mac GB2312
            case 81:
                return 10002; // Mac Big5
            case 82:
                return 0; // Mac Johab (old)
            case 83:
                return 10005; // Mac Hebrew
            case 84:
                return 10004; // Mac Arabic
            case 85:
                return 10006; // Mac Greek
            case 86:
                return 10081; // Mac Turkish
            case 87:
                return 10021; // Mac Thai
            case 88:
                return 10029; // Mac East Europe
            case 89:
                return 10007; // Mac Russian
            case 128:
                return 932; // Shift JIS
            case 129:
                return 949; // Hangul
            case 130:
                return 1361; // Johab
            case 134:
                return 936; // GB2312
            case 136:
                return 950; // Big5
            case 161:
                return 1253; // Greek
            case 162:
                return 1254; // Turkish
            case 163:
                return 1258; // Vietnamese
            case 177:
                return 1255; // Hebrew
            case 178:
                return 1256; // Arabic
            case 179:
                return 0; // Arabic Traditional (old)
            case 180:
                return 0; // Arabic user (old)
            case 181:
                return 0; // Hebrew user (old)
            case 186:
                return 1257; // Baltic
            case 204:
                return 1251; // Russian
            case 222:
                return 874; // Thai
            case 238:
                return 1250; // Eastern European
            case 254:
                return 437; // PC 437
            case 255:
                return 850; // OEM
        }

        return 0;
    }
    #endregion
}