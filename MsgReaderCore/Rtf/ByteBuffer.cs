//
// ColorTable.cs
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

using System;
using System.Text;

namespace MsgReader.Rtf;

/// <summary>
///     Binary data buffer
/// </summary>
internal class ByteBuffer
{
    #region Fields
    /// <summary>
    ///     Current contains validate bytes
    /// </summary>
    private int _intCount;

    /// <summary>
    ///     byte array
    /// </summary>
    private byte[] _buffer = new byte[16];
    #endregion

    #region Clear
    /// <summary>
    ///     Clear object's data
    /// </summary>
    public virtual void Clear()
    {
        _buffer = new byte[16];
        _intCount = 0;
    }
    #endregion

    #region Reset
    /// <summary>
    ///     Reset current position without clear buffer
    /// </summary>
    public void Reset()
    {
        _intCount = 0;
    }
    #endregion

    #region This
    /// <summary>
    ///     Set of get byte at special index which starts with 0
    /// </summary>
    public byte this[int index]
    {
        get => _buffer[index];
        set => _buffer[index] = value;
    }
    #endregion

    #region Count
    /// <summary>
    ///     Validate bytes count
    /// </summary>
    public virtual int Count => _intCount;
    #endregion

    #region Add
    /// <summary>
    ///     Add a byte
    /// </summary>
    /// <param name="b">byte data</param>
    public void Add(byte b)
    {
        FixBuffer(_intCount + 1);
        _buffer[_intCount] = b;
        _intCount++;
    }

    /// <summary>
    ///     Add bytes
    /// </summary>
    /// <param name="bs">bytes</param>
    public void Add(byte[] bs)
    {
        if (bs != null)
            Add(bs, 0, bs.Length);
    }

    /// <summary>
    ///     Add bytes
    /// </summary>
    /// <param name="bs">Bytes</param>
    /// <param name="startIndex">Start index</param>
    /// <param name="length">Length</param>
    public void Add(byte[] bs, int startIndex, int length)
    {
        if (bs != null && startIndex >= 0 && startIndex + length <= bs.Length && length > 0)
        {
            FixBuffer(_intCount + length);
            Array.Copy(bs, startIndex, _buffer, _intCount, length);
            _intCount += length;
        }
    }
    #endregion

    #region ToArray
    /// <summary>
    ///     Get validate bytes array
    /// </summary>
    /// <returns>bytes array</returns>
    public byte[] ToArray()
    {
        if (_intCount > 0)
        {
            var bs = new byte[_intCount];
            Array.Copy(_buffer, 0, bs, 0, _intCount);
            return bs;
        }

        return null;
    }
    #endregion

    #region GetString
    /// <summary>
    ///     Convert bytes data to a string
    /// </summary>
    /// <param name="encoding">string encoding</param>
    /// <returns>string data</returns>
    public string GetString(Encoding encoding)
    {
        if (encoding == null)
            throw new ArgumentNullException(nameof(encoding));
        return _intCount > 0 ? encoding.GetString(_buffer, 0, _intCount) : string.Empty;
    }
    #endregion

    #region FixBuffer
    /// <summary>
    ///     Fix inner buffer so it can fit to new size of buffer
    /// </summary>
    /// <param name="newSize">new size</param>
    protected void FixBuffer(int newSize)
    {
        if (newSize <= _buffer.Length)
            return;

        if (newSize < (int)(_buffer.Length * 1.5))
            newSize = (int)(_buffer.Length * 1.5);

        var bs = new byte[newSize];
        Buffer.BlockCopy(_buffer, 0, bs, 0, _buffer.Length);
        _buffer = bs;
    }
    #endregion
}