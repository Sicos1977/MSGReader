//
// TnefReaderStream.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
//
// Refactoring to the code done by Kees van Spelde so that it works in this project
// Copyright (c) 2023 Kees van Spelde <sicos2002@hotmail.com>
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
using System.IO;

namespace MsgReader.Tnef;

/// <summary>
///     A stream for reading raw values from a <see cref="MsgReader.Reader" /> or <see cref="PropertyReader" />.
/// </summary>
/// <remarks>
///     A stream for reading raw values from a <see cref="MsgReader.Reader" /> or <see cref="PropertyReader" />.
/// </remarks>
internal class ReaderStream : Stream
{
    #region Fields
    private readonly int _valueEndOffset;
    private readonly int _dataEndOffset;
    private readonly TnefReader _reader;
    private bool _disposed;
    #endregion

    #region Properties
    /// <summary>
    ///     Check whether or not the stream supports reading.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ReaderStream" /> is always readable.
    /// </remarks>
    /// <value><c>true</c> if the stream supports reading; otherwise, <c>false</c>.</value>
    public override bool CanRead => true;

    /// <summary>
    ///     Check whether or not the stream supports writing.
    /// </summary>
    /// <remarks>
    ///     Writing to a <see cref="ReaderStream" /> is not supported.
    /// </remarks>
    /// <value><c>true</c> if the stream supports writing; otherwise, <c>false</c>.</value>
    public override bool CanWrite => false;

    /// <summary>
    ///     Check whether or not the stream supports seeking.
    /// </summary>
    /// <remarks>
    ///     Seeking within a <see cref="ReaderStream" /> is not supported.
    /// </remarks>
    /// <value><c>true</c> if the stream supports seeking; otherwise, <c>false</c>.</value>
    public override bool CanSeek => false;

    /// <summary>
    ///     Get the length of the stream, in bytes.
    /// </summary>
    /// <remarks>
    ///     Getting the length of a <see cref="ReaderStream" /> is not supported.
    /// </remarks>
    /// <value>The length of the stream in bytes.</value>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support seeking.
    /// </exception>
    public override long Length => throw new NotSupportedException("Cannot get the length of the stream.");

    /// <summary>
    ///     Get or sets the current position within the stream.
    /// </summary>
    /// <remarks>
    ///     Getting and setting the position of a <see cref="ReaderStream" /> is not supported.
    /// </remarks>
    /// <value>The position of the stream.</value>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support seeking.
    /// </exception>
    public override long Position
    {
        get => throw new NotSupportedException("The stream does not support seeking.");
        set => throw new NotSupportedException("The stream does not support seeking.");
    }
    #endregion

    #region Constructor
    /// <summary>
    ///     Initialize a new instance of the <see cref="ReaderStream" /> class.
    /// </summary>
    /// <remarks>
    ///     Creates a stream for reading a raw value from the <see cref="MsgReader.Reader" />.
    /// </remarks>
    /// <param name="tnefReader">The <see cref="MsgReader.Reader" />.</param>
    /// <param name="dataEndOffset">The end offset of the data.</param>
    /// <param name="valueEndOffset">The end offset of the container value.</param>
    public ReaderStream(TnefReader tnefReader, int dataEndOffset, int valueEndOffset)
    {
        _valueEndOffset = valueEndOffset;
        _dataEndOffset = dataEndOffset;
        _reader = tnefReader;
    }
    #endregion

    #region ValidateArguments
    private static void ValidateArguments(byte[] buffer, int offset, int count)
    {
        if (buffer is null)
            throw new ArgumentNullException(nameof(buffer));

        if (offset < 0 || offset > buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(offset));

        if (count < 0 || count > buffer.Length - offset)
            throw new ArgumentOutOfRangeException(nameof(count));
    }
    #endregion

    #region Read
    /// <summary>
    ///     Read a sequence of bytes from the stream and advances the position
    ///     within the stream by the number of bytes read.
    /// </summary>
    /// <remarks>
    ///     Reads a sequence of bytes from the stream and advances the position
    ///     within the stream by the number of bytes read.
    /// </remarks>
    /// <returns>
    ///     The total number of bytes read into the buffer. This can be less than the number of bytes requested if that many
    ///     bytes are not currently available, or zero (0) if the end of the stream has been reached.
    /// </returns>
    /// <param name="buffer">The buffer to read data into.</param>
    /// <param name="offset">The offset into the buffer to start reading data.</param>
    /// <param name="count">The number of bytes to read.</param>
    /// <exception cref="ArgumentNullException">
    ///     <paramref name="buffer" /> is <c>null</c>.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    ///     <para><paramref name="offset" /> is less than zero or greater than the length of <paramref name="buffer" />.</para>
    ///     <para>-or-</para>
    ///     <para>
    ///         The <paramref name="buffer" /> is not large enough to contain <paramref name="count" /> bytes starting
    ///         at the specified <paramref name="offset" />.
    ///     </para>
    /// </exception>
    /// <exception cref="ObjectDisposedException">
    ///     The stream has been disposed.
    /// </exception>
    /// <exception cref="IOException">
    ///     An I/O error occurred.
    /// </exception>
    public override int Read(byte[] buffer, int offset, int count)
    {
        ValidateArguments(buffer, offset, count);

        CheckDisposed();

        var dataLeft = _dataEndOffset - _reader.StreamOffset;
        var n = Math.Min(dataLeft, count);

        var read = n > 0 ? _reader.ReadAttributeRawValue(buffer, offset, n) : 0;

        dataLeft -= read;

        if (dataLeft != 0 || _valueEndOffset <= _reader.StreamOffset) return read;
        var valueLeft = _valueEndOffset - _reader.StreamOffset;
        var buf = new byte[valueLeft];

        _reader.ReadAttributeRawValue(buf, 0, valueLeft);

        return read;
    }
    #endregion

    #region Write
    /// <summary>
    ///     Write a sequence of bytes to the stream and advances the current
    ///     position within this stream by the number of bytes written.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ReaderStream" /> does not support writing.
    /// </remarks>
    /// <param name="buffer">The buffer to write.</param>
    /// <param name="offset">The offset of the first byte to write.</param>
    /// <param name="count">The number of bytes to write.</param>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support writing.
    /// </exception>
    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotSupportedException("The stream does not support writing.");
    }
    #endregion

    #region Seek
    /// <summary>
    ///     Set the position within the current stream.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ReaderStream" /> does not support seeking.
    /// </remarks>
    /// <returns>The new position within the stream.</returns>
    /// <param name="offset">The offset into the stream relative to the <paramref name="origin" />.</param>
    /// <param name="origin">The origin to seek from.</param>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support seeking.
    /// </exception>
    public override long Seek(long offset, SeekOrigin origin)
    {
        throw new NotSupportedException("The stream does not support seeking.");
    }
    #endregion

    #region Flush
    /// <summary>
    ///     Clear all buffers for this stream and causes any buffered data to be written
    ///     to the underlying device.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ReaderStream" /> does not support writing.
    /// </remarks>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support writing.
    /// </exception>
    public override void Flush()
    {
        throw new NotSupportedException("The stream does not support writing.");
    }
    #endregion

    #region SetLength
    /// <summary>
    ///     Set the length of the stream.
    /// </summary>
    /// <remarks>
    ///     The <see cref="ReaderStream" /> does not support setting the length.
    /// </remarks>
    /// <param name="value">The desired length of the stream in bytes.</param>
    /// <exception cref="NotSupportedException">
    ///     The stream does not support setting the length.
    /// </exception>
    public override void SetLength(long value)
    {
        throw new NotSupportedException("The stream does not support setting the length.");
    }
    #endregion

    #region CheckDisposed
    private void CheckDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(ReaderStream));
    }
    #endregion

    #region Dispose
    /// <summary>
    ///     Release the unmanaged resources used by the <see cref="ReaderStream" /> and
    ///     optionally releases the managed resources.
    /// </summary>
    /// <remarks>
    ///     The underlying <see cref="MsgReader.Reader" /> is not disposed.
    /// </remarks>
    /// <param name="disposing">
    ///     <c>true</c> to release both managed and unmanaged resources;
    ///     <c>false</c> to release only the unmanaged resources.
    /// </param>
    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        _disposed = true;
    }
    #endregion
}