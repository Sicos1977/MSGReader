//
// TnefReader.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
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
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using MsgReader.Exceptions;
using MsgReader.Tnef.Enums;

// ReSharper disable UnusedMember.Global

namespace MsgReader.Tnef;

/// <summary>
///     A TNEF reader.
/// </summary>
/// <remarks>
///     A TNEF reader.
/// </remarks>
internal class TnefReader : IDisposable
{
    #region Consts
    internal const int TnefSignature = 0x223e9f78;
    private const int ReadAheadSize = 128;
    private const int BlockSize = 4096;
    private const int PadSize = 0;
    #endregion

    #region Fields
    // I/O buffering
    private readonly byte[] _input = new byte[ReadAheadSize + BlockSize + PadSize];
    private const int InputStart = ReadAheadSize;
    private int _inputIndex = ReadAheadSize;
    private int _inputEnd = ReadAheadSize;

    private long _position;
    private int _checksum;
    private int _codepage;
    private int _version;
    private bool _closed;
    private bool _eos;
    #endregion

    #region Properties
    /// <summary>
    ///     Get the attachment key value.
    /// </summary>
    /// <remarks>
    ///     Gets the attachment key value.
    /// </remarks>
    /// <value>The attachment key value.</value>
    public short AttachmentKey { get; private set; }

    /// <summary>
    ///     Get the current attribute's level.
    /// </summary>
    /// <remarks>
    ///     Gets the current attribute's level.
    /// </remarks>
    /// <value>The current attribute's level.</value>
    public AttributeLevel AttributeLevel { get; private set; }

    /// <summary>
    ///     Get the length of the current attribute's raw value.
    /// </summary>
    /// <remarks>
    ///     Gets the length of the current attribute's raw value.
    /// </remarks>
    /// <value>The length of the current attribute's raw value.</value>
    public int AttributeRawValueLength { get; private set; }

    /// <summary>
    ///     Get the stream offset of the current attribute's raw value.
    /// </summary>
    /// <remarks>
    ///     Gets the stream offset of the current attribute's raw value.
    /// </remarks>
    /// <value>The stream offset of the current attribute's raw value.</value>
    public int AttributeRawValueStreamOffset { get; private set; }

    /// <summary>
    ///     Get the current attribute's tag.
    /// </summary>
    /// <remarks>
    ///     Gets the current attribute's tag.
    /// </remarks>
    /// <value>The current attribute's tag.</value>
    public AttributeTag AttributeTag { get; private set; }

    internal AttributeType AttributeType => (AttributeType)((int)AttributeTag & 0xF0000);

    /// <summary>
    ///     Get the compliance mode.
    /// </summary>
    /// <remarks>
    ///     Gets the compliance mode.
    /// </remarks>
    /// <value>The compliance mode.</value>
    public ComplianceMode ComplianceMode { get; }

    /// <summary>
    ///     Get the current compliance status of the TNEF stream.
    /// </summary>
    /// <remarks>
    ///     <para>Gets the current compliance status of the TNEF stream.</para>
    ///     <para>As the reader progresses, this value may change if errors are encountered.</para>
    /// </remarks>
    /// <value>The compliance status.</value>
    public ComplianceStatus ComplianceStatus { get; internal set; }

    internal Stream InputStream { get; }

    /// <summary>
    ///     Get the message codepage.
    /// </summary>
    /// <remarks>
    ///     Gets the message codepage.
    /// </remarks>
    /// <value>The message codepage.</value>
    public int MessageCodepage
    {
        get => _codepage;
        private set
        {
            if (value == _codepage)
                return;

            try
            {
                var encoding = Encoding.GetEncoding(value);
                _codepage = encoding.CodePage;
            }
            catch (Exception ex)
            {
                ComplianceStatus |= ComplianceStatus.InvalidMessageCodepage;
                if (ComplianceMode == ComplianceMode.Strict)
                    throw new MRTnefException(ComplianceStatus.InvalidMessageCodepage,
                        string.Format(CultureInfo.InvariantCulture, "Invalid message codepage: {0}", value), ex);
                _codepage = 1252;
            }
        }
    }

    /// <summary>
    ///     Get the TNEF property reader.
    /// </summary>
    /// <remarks>
    ///     Gets the TNEF property reader.
    /// </remarks>
    /// <value>The TNEF property reader.</value>
    public PropertyReader TnefPropertyReader { get; }

    /// <summary>
    ///     Get the current stream offset.
    /// </summary>
    /// <remarks>
    ///     Gets the current stream offset.
    /// </remarks>
    /// <value>The stream offset.</value>
    public int StreamOffset => (int)(_position - (_inputEnd - _inputIndex));

    /// <summary>
    ///     Get the TNEF version.
    /// </summary>
    /// <remarks>
    ///     Gets the TNEF version.
    /// </remarks>
    /// <value>The TNEF version.</value>
    public int TnefVersion
    {
        get => _version;
        private set
        {
            if (value != 0x00010000)
            {
                ComplianceStatus |= ComplianceStatus.InvalidTnefVersion;
                if (ComplianceMode == ComplianceMode.Strict)
                    throw new MRTnefException(ComplianceStatus.InvalidTnefVersion,
                        string.Format(CultureInfo.InvariantCulture, "Invalid TNEF version: {0}", value));
            }

            _version = value;
        }
    }
    #endregion

    #region Constructors
    /// <summary>
    ///     Initialize a new instance of the <see cref="TnefReader" /> class.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         When reading a TNEF stream using the <see cref="Enums.ComplianceMode.Strict" /> mode,
    ///         a <see cref="MRTnefException" /> will be thrown immediately at the first sign of
    ///         invalid or corrupted data.
    ///     </para>
    ///     <para>
    ///         When reading a TNEF stream using the <see cref="Enums.ComplianceMode.Loose" /> mode,
    ///         however, compliance issues are accumulated in the <see cref="ComplianceMode" />
    ///         property, but exceptions are not raised unless the stream is too corrupted to continue.
    ///     </para>
    /// </remarks>
    /// <param name="inputStream">The input stream.</param>
    /// <param name="defaultMessageCodepage">The default message codepage.</param>
    /// <param name="complianceMode">The compliance mode.</param>
    /// <exception cref="ArgumentNullException">
    ///     <paramref name="inputStream" /> is <c>null</c>.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    ///     <paramref name="defaultMessageCodepage" /> is not a valid codepage.
    /// </exception>
    /// <exception cref="NotSupportedException">
    ///     <paramref name="defaultMessageCodepage" /> is not a supported codepage.
    /// </exception>
    /// <exception cref="MRTnefException">
    ///     The TNEF stream is corrupted or invalid.
    /// </exception>
    public TnefReader(Stream inputStream, int defaultMessageCodepage, ComplianceMode complianceMode)
    {
        if (inputStream is null)
            throw new ArgumentNullException(nameof(inputStream));

        if (defaultMessageCodepage < 0)
            throw new ArgumentOutOfRangeException(nameof(defaultMessageCodepage));

        if (defaultMessageCodepage != 0)
        {
            // make sure that this codepage is valid...
            var encoding = Encoding.GetEncoding(defaultMessageCodepage);
            _codepage = encoding.CodePage;
        }
        else
        {
            _codepage = 1252;
        }

        TnefPropertyReader = new PropertyReader(this);
        ComplianceMode = complianceMode;
        InputStream = inputStream;

        DecodeHeader();
    }

    /// <summary>
    ///     Initialize a new instance of the <see cref="TnefReader" /> class.
    /// </summary>
    /// <remarks>
    ///     Creates a new <see cref="TnefReader" /> for the specified input stream.
    /// </remarks>
    /// <param name="inputStream">The input stream.</param>
    /// <exception cref="ArgumentNullException">
    ///     <paramref name="inputStream" /> is <c>null</c>.
    /// </exception>
    public TnefReader(Stream inputStream) : this(inputStream, 0, ComplianceMode.Loose)
    {
    }
    #endregion

    #region Destructor
    /// <summary>
    ///     Releases unmanaged resources and performs other cleanup operations before the
    ///     <see cref="TnefReader" /> is reclaimed by garbage collection.
    /// </summary>
    /// <remarks>
    ///     Releases unmanaged resources and performs other cleanup operations before the
    ///     <see cref="TnefReader" /> is reclaimed by garbage collection.
    /// </remarks>
    ~TnefReader()
    {
        Dispose(false);
    }
    #endregion

    #region CheckDisposed
    private void CheckDisposed()
    {
        if (_closed)
            throw new ObjectDisposedException("TnefReader");
    }
    #endregion

    #region ReadAhead
    internal int ReadAhead(int atLeast)
    {
        CheckDisposed();

        var left = _inputEnd - _inputIndex;

        if (left >= atLeast || _eos)
            return left;

        var index = _inputIndex;
        var start = InputStart;
        var end = _inputEnd;
        int read;

        // attempt to align the end of the remaining input with ReadAheadSize
        if (index >= start)
        {
            start -= Math.Min(ReadAheadSize, left);
            Buffer.BlockCopy(_input, index, _input, start, left);
            index = start;
            start += left;
        }
        else if (index > 0)
        {
            var shift = Math.Min(index, end - start);
            Buffer.BlockCopy(_input, index, _input, index - shift, left);
            index -= shift;
            start = index + left;
        }
        else
        {
            // we can't shift...
            start = end;
        }

        _inputIndex = index;
        _inputEnd = start;

        end = _input.Length - PadSize;

        if ((read = InputStream.Read(_input, start, end - start)) > 0)
        {
            _inputEnd += read;
            _position += read;
        }
        else
        {
            _eos = true;
        }

        return _inputEnd - _inputIndex;
    }
    #endregion

    #region SetComplianceError
    internal void SetComplianceError(ComplianceStatus error, Exception innerException = null)
    {
        ComplianceStatus |= error;

        if (ComplianceMode != ComplianceMode.Strict)
            return;

        string message;

        switch (error)
        {
            case ComplianceStatus.AttributeOverflow:
                message = "Too many attributes.";
                break;
            case ComplianceStatus.InvalidAttribute:
                message = "Invalid attribute.";
                break;
            case ComplianceStatus.InvalidAttributeChecksum:
                message = "Invalid attribute checksum.";
                break;
            case ComplianceStatus.InvalidAttributeLength:
                message = "Invalid attribute length.";
                break;
            case ComplianceStatus.InvalidAttributeLevel:
                message = "Invalid attribute level.";
                break;
            case ComplianceStatus.InvalidAttributeValue:
                message = "Invalid attribute value.";
                break;
            case ComplianceStatus.InvalidDate:
                message = "Invalid date.";
                break;
            case ComplianceStatus.InvalidMessageClass:
                message = "Invalid message class.";
                break;
            case ComplianceStatus.InvalidMessageCodepage:
                message = "Invalid message codepage.";
                break;
            case ComplianceStatus.InvalidPropertyLength:
                message = "Invalid property length.";
                break;
            case ComplianceStatus.InvalidRowCount:
                message = "Invalid row count.";
                break;
            case ComplianceStatus.InvalidTnefSignature:
                message = "Invalid TNEF signature.";
                break;
            case ComplianceStatus.InvalidTnefVersion:
                message = "Invalid TNEF version.";
                break;
            case ComplianceStatus.NestingTooDeep:
                message = "Nesting too deep.";
                break;
            case ComplianceStatus.StreamTruncated:
                message = "Truncated TNEF stream.";
                break;
            case ComplianceStatus.UnsupportedPropertyType:
                message = "Unsupported property type.";
                break;
            case ComplianceStatus.Compliant:
                return;

            default:
                throw new ArgumentOutOfRangeException(nameof(error), error, null);
        }

        if (innerException != null)
            throw new MRTnefException(error, message, innerException);

        throw new MRTnefException(error, message);
    }
    #endregion

    #region DecodeHeader
    private void DecodeHeader()
    {
        try
        {
            // read the TNEFSignature
            var signature = ReadInt32();
            if (signature != TnefSignature)
                SetComplianceError(ComplianceStatus.InvalidTnefSignature);

            // read the LegacyKey (ignore this value)
            AttachmentKey = ReadInt16();
        }
        catch (EndOfStreamException)
        {
            SetComplianceError(ComplianceStatus.StreamTruncated);
        }
    }
    #endregion

    #region CheckAttributeLevel
    private void CheckAttributeLevel()
    {
        switch (AttributeLevel)
        {
            case AttributeLevel.Attachment:
            case AttributeLevel.Message:
                break;
            default:
                SetComplianceError(ComplianceStatus.InvalidAttributeLevel);
                break;
        }
    }
    #endregion

    #region CheckAttributeTag
    private void CheckAttributeTag()
    {
        switch (AttributeTag)
        {
            case AttributeTag.AidOwner:
            case AttributeTag.AttachCreateDate:
            case AttributeTag.AttachData:
            case AttributeTag.Attachment:
            case AttributeTag.AttachMetaFile:
            case AttributeTag.AttachModifyDate:
            case AttributeTag.AttachTitle:
            case AttributeTag.AttachTransportFilename:
            case AttributeTag.Body:
            case AttributeTag.ConversationId:
            case AttributeTag.DateEnd:
            case AttributeTag.DateModified:
            case AttributeTag.DateReceived:
            case AttributeTag.DateSent:
            case AttributeTag.DateStart:
            case AttributeTag.Delegate:
            case AttributeTag.From:
            case AttributeTag.MapiProperties:
            case AttributeTag.MessageClass:
            case AttributeTag.MessageId:
            case AttributeTag.MessageStatus:
            case AttributeTag.Null:
            case AttributeTag.OriginalMessageClass:
            case AttributeTag.Owner:
            case AttributeTag.ParentId:
            case AttributeTag.Priority:
            case AttributeTag.RecipientTable:
            case AttributeTag.RequestResponse:
            case AttributeTag.SentFor:
            case AttributeTag.Subject:
                break;
            case AttributeTag.AttachRenderData:
                TnefPropertyReader.AttachMethod = AttachMethod.ByValue;
                break;
            case AttributeTag.OemCodepage:
                MessageCodepage = PeekInt32();
                break;
            case AttributeTag.TnefVersion:
                TnefVersion = PeekInt32();
                break;
            default:
                SetComplianceError(ComplianceStatus.InvalidAttribute);
                break;
        }
    }
    #endregion

    #region Read
    internal byte ReadByte()
    {
        if (ReadAhead(1) < 1)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 1);

        return _input[_inputIndex++];
    }

    internal short ReadInt16()
    {
        if (ReadAhead(2) < 2)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 2);

        var result = BinaryPrimitives.ReadInt16LittleEndian(_input.AsSpan(_inputIndex));

        _inputIndex += 2;

        return result;
    }

    internal int ReadInt32()
    {
        if (ReadAhead(4) < 4)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 4);

        var result = BinaryPrimitives.ReadInt32LittleEndian(_input.AsSpan(_inputIndex));

        _inputIndex += 4;

        return result;
    }

    internal long ReadInt64()
    {
        if (ReadAhead(8) < 8)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 8);

        var result = BinaryPrimitives.ReadInt64LittleEndian(_input.AsSpan(_inputIndex));

        _inputIndex += 8;

        return result;
    }
    #endregion

    #region PeekInt32
    internal int PeekInt32()
    {
        if (ReadAhead(4) < 4)
            throw new EndOfStreamException();

        return BinaryPrimitives.ReadInt32LittleEndian(_input.AsSpan(_inputIndex));
    }
    #endregion

    #region ReadSingle
    internal float ReadSingle()
    {
        if (ReadAhead(4) < 4)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 4);
        
        var result = BitConverter.ToSingle(_input, _inputIndex);
        
        _inputIndex += 4;

        return result;
    }
    #endregion

    #region ReadDouble
    internal double ReadDouble()
    {
        if (ReadAhead(8) < 8)
            throw new EndOfStreamException();

        UpdateChecksum(_input, _inputIndex, 8);

        var result = BitConverter.ToDouble(_input, _inputIndex);

        _inputIndex += 8;

        return result;
    }
    #endregion

    #region Seek
    internal bool Seek(int offset)
    {
        CheckDisposed();

        var left = offset - StreamOffset;

        if (left <= 0)
            return true;

        do
        {
            var n = Math.Min(_inputEnd - _inputIndex, left);

            UpdateChecksum(_input, _inputIndex, n);
            _inputIndex += n;
            left -= n;

            if (left == 0)
                break;

            if (ReadAhead(left) != 0) continue;
            SetComplianceError(ComplianceStatus.StreamTruncated);
            return false;
        } while (true);

        return true;
    }
    #endregion

    #region SkipAttributeRawValue
    private bool SkipAttributeRawValue()
    {
        var offset = AttributeRawValueStreamOffset + AttributeRawValueLength;
        int actual;

        if (!Seek(offset))
            return false;

        // Note: ReadInt16() will update the checksum, so we need to capture it here
        var expected = _checksum;

        try
        {
            actual = (ushort)ReadInt16();
        }
        catch (EndOfStreamException)
        {
            SetComplianceError(ComplianceStatus.StreamTruncated);
            return false;
        }

        if (actual != expected)
            SetComplianceError(ComplianceStatus.InvalidAttributeChecksum);

        return true;
    }
    #endregion

    #region ReadNextAttribute
    /// <summary>
    ///     Advance to the next attribute in the TNEF stream.
    /// </summary>
    /// <remarks>
    ///     Advances to the next attribute in the TNEF stream.
    /// </remarks>
    /// <returns><c>true</c> if there is another attribute available to be read; otherwise <c>false</c>.</returns>
    /// <exception cref="MRTnefException">
    ///     The TNEF stream is corrupted or invalid.
    /// </exception>
    public bool ReadNextAttribute()
    {
        CheckDisposed();

        if (AttributeRawValueStreamOffset != 0 && !SkipAttributeRawValue())
            return false;

        try
        {
            AttributeLevel = (AttributeLevel)ReadByte();
        }
        catch (EndOfStreamException)
        {
            return false;
        }

        CheckAttributeLevel();

        try
        {
            AttributeTag = (AttributeTag)ReadInt32();
            AttributeRawValueLength = ReadInt32();
            AttributeRawValueStreamOffset = StreamOffset;
            _checksum = 0;
        }
        catch (EndOfStreamException)
        {
            SetComplianceError(ComplianceStatus.StreamTruncated);
            return false;
        }

        CheckAttributeTag();

        if (AttributeRawValueLength < 0)
        {
            SetComplianceError(ComplianceStatus.InvalidAttributeLength);
            return false;
        }

        try
        {
            TnefPropertyReader.Load();
        }
        catch (EndOfStreamException)
        {
            SetComplianceError(ComplianceStatus.StreamTruncated);
            return false;
        }

        return true;
    }
    #endregion

    #region UpdateChecksum
    private void UpdateChecksum(IReadOnlyList<byte> buffer, int offset, int count)
    {
        for (var i = offset; i < offset + count; i++)
            _checksum = (_checksum + buffer[i]) & 0xFFFF;
    }
    #endregion

    #region ReadAttributeRawValue
    /// <summary>
    ///     Read the raw attribute value data from the underlying TNEF stream.
    /// </summary>
    /// <remarks>
    ///     Reads the raw attribute value data from the underlying TNEF stream.
    /// </remarks>
    /// <returns>
    ///     The total number of bytes read into the buffer. This can be less than the number
    ///     of bytes requested if that many bytes are not available, or zero (0) if the end of the
    ///     value has been reached.
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
    public int ReadAttributeRawValue(byte[] buffer, int offset, int count)
    {
        if (buffer is null)
            throw new ArgumentNullException(nameof(buffer));

        if (offset < 0 || offset >= buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(offset));

        if (count < 0 || count > buffer.Length - offset)
            throw new ArgumentOutOfRangeException(nameof(count));

        CheckDisposed();

        var dataEndOffset = AttributeRawValueStreamOffset + AttributeRawValueLength;
        var dataLeft = dataEndOffset - StreamOffset;

        if (dataLeft == 0)
            return 0;

        var inputLeft = _inputEnd - _inputIndex;
        var n = Math.Min(dataLeft, count);

        if (n > inputLeft && inputLeft < ReadAheadSize)
        {
            if ((n = Math.Min(ReadAhead(n), n)) == 0)
            {
                SetComplianceError(ComplianceStatus.StreamTruncated);
                return 0;
            }
        }
        else
        {
            n = Math.Min(inputLeft, n);
        }

        Buffer.BlockCopy(_input, _inputIndex, buffer, offset, n);
        UpdateChecksum(buffer, offset, n);
        _inputIndex += n;

        return n;
    }
    #endregion

    #region ResetComplianceStatus
    /// <summary>
    ///     Reset the compliance status.
    /// </summary>
    /// <remarks>
    ///     Resets the compliance status.
    /// </remarks>
    public void ResetComplianceStatus()
    {
        ComplianceStatus = ComplianceStatus.Compliant;
    }
    #endregion

    #region Close
    /// <summary>
    ///     Close the TNEF reader and the underlying stream.
    /// </summary>
    /// <remarks>
    ///     Closes the TNEF reader and the underlying stream.
    /// </remarks>
    public void Close()
    {
        Dispose();
    }
    #endregion

    #region IDisposable implementation
    /// <summary>
    ///     Release the unmanaged resources used by the <see cref="TnefReader" /> and
    ///     optionally releases the managed resources.
    /// </summary>
    /// <remarks>
    ///     Releases the unmanaged resources used by the <see cref="TnefReader" /> and
    ///     optionally releases the managed resources.
    /// </remarks>
    /// <param name="disposing">
    ///     <c>true</c> to release both managed and unmanaged resources;
    ///     <c>false</c> to release only the unmanaged resources.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (disposing && !_closed)
            InputStream.Dispose();
    }

    /// <summary>
    ///     Release all resource used by the <see cref="TnefReader" /> object.
    /// </summary>
    /// <remarks>
    ///     Call <see cref="Dispose()" /> when you are finished using the <see cref="TnefReader" />. The
    ///     <see cref="Dispose()" /> method leaves the <see cref="TnefReader" /> in an unusable state. After calling
    ///     <see cref="Dispose()" />, you must release all references to the <see cref="TnefReader" /> so the garbage
    ///     collector can reclaim the memory that the <see cref="TnefReader" /> was occupying.
    /// </remarks>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
        _closed = true;
    }
    #endregion
}