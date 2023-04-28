//
// FontTable.cs
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

using System;
using System.Collections.Generic;
using System.Text;

namespace MsgReader.Rtf;

/// <summary>
///     Rtf font information
/// </summary>
internal class Font
{
    #region Fields
    private static Dictionary<int, Encoding> _encodingCharsets;
    private int _charset = 1;
    #endregion

    #region Properties
    /// <summary>
    ///     Font index
    /// </summary>
    public int Index { get; set; }

    public bool NilFlag { get; set; }

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
            Encoding = GetRtfEncoding(_charset);
        }
    }

    /// <summary>
    ///     Encoding
    /// </summary>
    public Encoding Encoding { get; private set; }
    #endregion

    #region Constructor
    /// <summary>
    ///     Initialize instance
    /// </summary>
    /// <param name="index">font index</param>
    /// <param name="name">font name</param>
    public Font(int index, string name)
    {
        NilFlag = false;
        Encoding = null;
        Index = index;
        Name = name;
    }
    #endregion

    #region Clone
    public Font Clone()
    {
        return new Font(Index, Name)
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
        return Index + ":" + Name + " Charset:" + _charset;
    }
    #endregion

    #region CheckEncodingCharsets
    private static void CheckEncodingCharsets()
    {
        if (_encodingCharsets == null)
        {
            _encodingCharsets = new Dictionary<int, Encoding>();
            _encodingCharsets[77] = Encoding.GetEncoding(10000); // Mac Roman
            _encodingCharsets[78] = Encoding.GetEncoding(10001); // Mac Shift Jis
            _encodingCharsets[79] = Encoding.GetEncoding(10003); // Mac Hangul
            _encodingCharsets[80] = Encoding.GetEncoding(10008); // Mac GB2312
            _encodingCharsets[81] = Encoding.GetEncoding(10002); // Mac Big5
            _encodingCharsets[83] = Encoding.GetEncoding(10005); // Mac Hebrew
            _encodingCharsets[84] = Encoding.GetEncoding(10004); // Mac Arabic
            _encodingCharsets[85] = Encoding.GetEncoding(10006); // Mac Greek
            _encodingCharsets[86] = Encoding.GetEncoding(10081); // Mac Turkish
            _encodingCharsets[87] = Encoding.GetEncoding(10021); // Mac Thai
            _encodingCharsets[88] = Encoding.GetEncoding(10029); // Mac East Europe
            _encodingCharsets[89] = Encoding.GetEncoding(10007); // Mac Russian
            _encodingCharsets[128] = Encoding.GetEncoding(932); // Shift JIS
            _encodingCharsets[129] = Encoding.GetEncoding(949); // Hangul
            _encodingCharsets[130] = Encoding.GetEncoding(1361); // Johab
            _encodingCharsets[134] = Encoding.GetEncoding(936); // GB2312
            _encodingCharsets[136] = Encoding.GetEncoding(950); // Big5
            _encodingCharsets[161] = Encoding.GetEncoding(1253); // Greek
            _encodingCharsets[162] = Encoding.GetEncoding(1254); // Turkish
            _encodingCharsets[163] = Encoding.GetEncoding(1258); // Vietnamese
            _encodingCharsets[177] = Encoding.GetEncoding(1255); // Hebrew
            _encodingCharsets[178] = Encoding.GetEncoding(1256); // Arabic 
            _encodingCharsets[186] = Encoding.GetEncoding(1257); // Baltic
            _encodingCharsets[204] = Encoding.GetEncoding(1251); // Russian
            _encodingCharsets[222] = Encoding.GetEncoding(874); // Thai
            _encodingCharsets[238] = Encoding.GetEncoding(1250); // Eastern European
            _encodingCharsets[254] = Encoding.GetEncoding(437); // PC 437
            _encodingCharsets[255] = Encoding.GetEncoding(850); // OEM
        }
    }
    #endregion

    #region GetCharset
    internal static int GetCharset(Encoding encoding)
    {
        CheckEncodingCharsets();
        foreach (var key in _encodingCharsets.Keys)
            // ReSharper disable once PossibleUnintendedReferenceComparison
            if (_encodingCharsets[key] == encoding)
                return key;

        return 1;
    }
    #endregion

    #region GetRTFEncoding
    internal static Encoding GetRtfEncoding(int chartset)
    {
        switch (chartset)
        {
            case 0:
                return AnsiEncoding.Instance;
            case 1:
                return Encoding.Default;
            default:
                CheckEncodingCharsets();
                return _encodingCharsets.ContainsKey(chartset) ? _encodingCharsets[chartset] : null;
        }
    }
    #endregion
}

#region Internal class ANSIEncoding
/// <summary>
///     Internal encoding for ansi
/// </summary>
internal class AnsiEncoding : Encoding
{
    public static AnsiEncoding Instance = new();

    public override string GetString(byte[] bytes, int index, int count)
    {
        var stringBuilder = new StringBuilder();
        var endIndex = Math.Min(bytes.Length - 1, index + count - 1);

        for (var iCount = index; iCount <= endIndex; iCount++)
            stringBuilder.Append(System.Convert.ToChar(bytes[iCount]));

        return stringBuilder.ToString();
    }

    public override int GetByteCount(char[] chars, int index, int count)
    {
        throw new Exception("The method or operation is not implemented.");
    }

    public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex)
    {
        throw new Exception("The method or operation is not implemented.");
    }

    public override int GetCharCount(byte[] bytes, int index, int count)
    {
        throw new Exception("The method or operation is not implemented.");
    }

    public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex)
    {
        throw new Exception("The method or operation is not implemented.");
    }

    public override int GetMaxByteCount(int charCount)
    {
        throw new Exception("The method or operation is not implemented.");
    }

    public override int GetMaxCharCount(int byteCount)
    {
        throw new Exception("The method or operation is not implemented.");
    }
}
#endregion