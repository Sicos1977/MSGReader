//
// TnefPropertyReader.cs
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
using System.Text;
using MsgReader.Exceptions;
using MsgReader.Tnef.Enums;

// ReSharper disable UnusedMember.Global

namespace MsgReader.Tnef;

/// <summary>
///     A TNEF property reader.
/// </summary>
/// <remarks>
///     A TNEF property reader.
/// </remarks>
internal class PropertyReader
{
    private static readonly Encoding DefaultEncoding = Encoding.GetEncoding(1252);

    #region Consts
    // Note: these constants taken from Microsoft's Reference Source in DateTime.cs
    private const long TicksPerMillisecond = 10000;
    private const long TicksPerSecond = TicksPerMillisecond * 1000;
    private const long TicksPerMinute = TicksPerSecond * 60;
    private const long TicksPerHour = TicksPerMinute * 60;
    private const long TicksPerDay = TicksPerHour * 24;

    private const int MillisPerSecond = 1000;
    private const int MillisPerMinute = MillisPerSecond * 60;
    private const int MillisPerHour = MillisPerMinute * 60;
    private const int MillisPerDay = MillisPerHour * 24;

    private const int DaysPerYear = 365;
    private const int DaysPer4Years = DaysPerYear * 4 + 1;
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;
    private const int DaysPer400Years = DaysPer100Years * 4 + 1;
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    private const int DaysTo10000 = DaysPer400Years * 25 - 366;

    private const long MaxMillis = (long)DaysTo10000 * MillisPerDay;

    private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;
    // private const long OaDateMinAsTicks = (DaysPer100Years - DaysPerYear) * TicksPerDay;
    private const double OaDateMinAsDouble = -657435.0;
    private const double OaDateMaxAsDouble = 2958466.0;
    #endregion

    #region Fields
    private PropertyTag _propertyTag;
    private readonly TnefReader _reader;
    private int _rawValueLength;
    private int _propertyIndex;
    private Decoder _decoder;
    private int _valueIndex;
    private int _rowIndex;
    #endregion

    #region Properties
    internal AttachMethod AttachMethod { get; set; }

    /// <summary>
    ///     Get a value indicating whether the current property is an embedded TNEF message.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether the current property is an embedded TNEF message.
    /// </remarks>
    /// <value><c>true</c> if the current property is an embedded TNEF message; otherwise, <c>false</c>.</value>
    public bool IsEmbeddedMessage => _propertyTag.Id == PropertyId.AttachData && AttachMethod == AttachMethod.EmbeddedMessage;


    /// <summary>
    ///     Get a value indicating whether or not the current property has multiple values.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether or not the current property has multiple values.
    /// </remarks>
    /// <value><c>true</c> if the current property has multiple values; otherwise, <c>false</c>.</value>
    public bool IsMultiValuedProperty => _propertyTag.IsMultiValued;

    /// <summary>
    ///     Get a value indicating whether or not the current property is a named property.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether or not the current property is a named property.
    /// </remarks>
    /// <value><c>true</c> if the current property is a named property; otherwise, <c>false</c>.</value>
    public bool IsNamedProperty => _propertyTag.IsNamed;

    /// <summary>
    ///     Get a value indicating whether the current property contains object values.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether the current property contains object values.
    /// </remarks>
    /// <value><c>true</c> if the current property contains object values; otherwise, <c>false</c>.</value>
    public bool IsObjectProperty => _propertyTag.ValueTnefType == PropertyType.Object;

    /// <summary>
    ///     Get the number of properties available.
    /// </summary>
    /// <remarks>
    ///     Gets the number of properties available.
    /// </remarks>
    /// <value>The property count.</value>
    public int PropertyCount { get; private set; }

    /// <summary>
    ///     Get the property name identifier.
    /// </summary>
    /// <remarks>
    ///     Gets the property name identifier.
    /// </remarks>
    /// <value>The property name identifier.</value>
    public NameId PropertyNameId { get; private set; }

    /// <summary>
    ///     Get the property tag.
    /// </summary>
    /// <remarks>
    ///     Gets the property tag.
    /// </remarks>
    /// <value>The property tag.</value>
    public PropertyTag PropertyTag => _propertyTag;

    /// <summary>
    ///     Get the length of the raw value.
    /// </summary>
    /// <remarks>
    ///     Gets the length of the raw value.
    /// </remarks>
    /// <value>The length of the raw value.</value>
    public int RawValueLength => _rawValueLength;

    /// <summary>
    ///     Get the raw value stream offset.
    /// </summary>
    /// <remarks>
    ///     Gets the raw value stream offset.
    /// </remarks>
    /// <value>The raw value stream offset.</value>
    public int RawValueStreamOffset { get; private set; }

    /// <summary>
    ///     Get the number of table rows available.
    /// </summary>
    /// <remarks>
    ///     Gets the number of table rows available.
    /// </remarks>
    /// <value>The row count.</value>
    public int RowCount { get; private set; }

    /// <summary>
    ///     Get the number of values available.
    /// </summary>
    /// <remarks>
    ///     Gets the number of values available.
    /// </remarks>
    /// <value>The value count.</value>
    public int ValueCount { get; private set; }

    /// <summary>
    ///     Get the type of the value.
    /// </summary>
    /// <remarks>
    ///     Gets the type of the value.
    /// </remarks>
    /// <value>The type of the value.</value>
    public Type ValueType
    {
        get
        {
            if (PropertyCount > 0)
                return GetPropertyValueType();

            return GetAttributeValueType();
        }
    }
    #endregion

    #region Constructor
    internal PropertyReader(TnefReader tnef)
    {
        _propertyTag = PropertyTag.Null;
        PropertyNameId = new NameId();
        RawValueStreamOffset = 0;
        _rawValueLength = 0;
        _propertyIndex = 0;
        PropertyCount = 0;
        _valueIndex = 0;
        ValueCount = 0;
        _rowIndex = 0;
        RowCount = 0;

        _reader = tnef;
    }
    #endregion

    #region GetEmbeddedMessageReader
    /// <summary>
    ///     Get the embedded TNEF message reader.
    /// </summary>
    /// <remarks>
    ///     Gets the embedded TNEF message reader.
    /// </remarks>
    /// <returns>The embedded TNEF message reader.</returns>
    /// <exception cref="InvalidOperationException">
    ///     <para>The property does not contain any more values.</para>
    ///     <para>-or-</para>
    ///     <para>The property value is not an embedded message.</para>
    /// </exception>
    public TnefReader GetEmbeddedMessageReader()
    {
        if (!IsEmbeddedMessage)
            throw new InvalidOperationException();

        var stream = GetRawValueReadStream();
        var guid = new byte[16];

        // ReSharper disable once MustUseReturnValue
        stream.Read(guid, 0, 16);

        return new TnefReader(stream, _reader.MessageCodepage, _reader.ComplianceMode);
    }
    #endregion

    #region GetRawValueReadStream
    /// <summary>
    ///     Get the raw value of the attribute or property as a stream.
    /// </summary>
    /// <remarks>
    ///     Gets the raw value of the attribute or property as a stream.
    /// </remarks>
    /// <returns>The raw value stream.</returns>
    /// <exception cref="InvalidOperationException">
    ///     The property does not contain any more values.
    /// </exception>
    public Stream GetRawValueReadStream()
    {
        if (_valueIndex >= ValueCount)
            throw new InvalidOperationException();

        var startOffset = RawValueStreamOffset;
        var length = RawValueLength;

        if (PropertyCount > 0 && _reader.StreamOffset == RawValueStreamOffset)
            switch (_propertyTag.ValueTnefType)
            {
                case PropertyType.Unicode:
                case PropertyType.String8:
                case PropertyType.Binary:
                case PropertyType.Object:
                    var n = ReadInt32();
                    if (n >= 0 && n + 4 < length)
                        length = n + 4;
                    break;
            }

        _valueIndex++;

        var valueEndOffset = startOffset + RawValueLength;
        var dataEndOffset = startOffset + length;

        return new ReaderStream(_reader, dataEndOffset, valueEndOffset);
    }
    #endregion

    #region CheckRawValueLength
    private bool CheckRawValueLength()
    {
        // Check that the property value does not go beyond the end of the end of the attribute
        var attrEndOffset = _reader.AttributeRawValueStreamOffset + _reader.AttributeRawValueLength;
        var valueEndOffset = RawValueStreamOffset + RawValueLength;

        if (valueEndOffset > attrEndOffset)
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidAttributeValue);
            return false;
        }

        return true;
    }
    #endregion

    #region ReadByte
    private byte ReadByte()
    {
        return _reader.ReadByte();
    }
    #endregion

    #region ReadBytes
    private byte[] ReadBytes(int count)
    {
        var bytes = new byte[count];
        var offset = 0;
        int read;

        while (offset < count && (read = _reader.ReadAttributeRawValue(bytes, offset, count - offset)) > 0)
            offset += read;

        return bytes;
    }
    #endregion

    #region ReadInt16
    private short ReadInt16()
    {
        return _reader.ReadInt16();
    }
    #endregion

    #region ReadInt32
    private int ReadInt32()
    {
        return _reader.ReadInt32();
    }
    #endregion

    #region PeekInt32
    private int PeekInt32()
    {
        return _reader.PeekInt32();
    }
    #endregion

    #region ReadInt64
    private long ReadInt64()
    {
        return _reader.ReadInt64();
    }
    #endregion

    #region ReadSingle
    private float ReadSingle()
    {
        return _reader.ReadSingle();
    }
    #endregion

    #region ReadDouble
    private double ReadDouble()
    {
        return _reader.ReadDouble();
    }
    #endregion

    #region DoubleDateToTicks
    // Note: this method taken from Microsoft's Reference Source in DateTime.cs
    private static long DoubleDateToTicks(double value)
    {
        // The check done this way will take care of NaN
        if (!(value < OaDateMaxAsDouble) || !(value > OaDateMinAsDouble))
            throw new ArgumentException(@"Invalid OLE Automation Date.", nameof(value));

        var millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

        if (millis < 0)
            millis -= millis % MillisPerDay * 2;

        millis += DoubleDateOffset / TicksPerMillisecond;

        if (millis is < 0 or >= MaxMillis)
            throw new ArgumentException(@"Invalid OLE Automation Date.", nameof(value));

        return millis * TicksPerMillisecond;
    }
    #endregion

    #region ReadAppTime
    private DateTime ReadAppTime()
    {
        var appTime = ReadDouble();

        // Note: equivalent to DateTime.FromOADate(). Unfortunately, FromOADate() is
        // not available in some PCL profiles.
        return new DateTime(DoubleDateToTicks(appTime), DateTimeKind.Unspecified);
    }
    #endregion

    #region ReadSysTime
    private DateTime ReadSysTime()
    {
        var fileTime = ReadInt64();

        return DateTime.FromFileTime(fileTime);
    }
    #endregion

    #region MyRegion
    private static int GetPaddedLength(int length)
    {
        return (length + 3) & ~3;
    }
    #endregion

    #region ReadByteArray
    private byte[] ReadByteArray()
    {
        var length = ReadInt32();
        var bytes = ReadBytes(length);

        if (length % 4 != 0)
        {
            // remaining bytes are padding
            var padding = 4 - length % 4;

            _reader.Seek(_reader.StreamOffset + padding);
        }

        return bytes;
    }
    #endregion

    #region ReadUnicodeString
    private string ReadUnicodeString()
    {
        var bytes = ReadByteArray();
        var length = bytes.Length;

        // force length to a multiple of 2 bytes
        length &= ~1;

        while (length > 1 && bytes[length - 1] == 0 && bytes[length - 2] == 0)
            length -= 2;

        return length < 2 ? string.Empty : Encoding.Unicode.GetString(bytes, 0, length);
    }
    #endregion

    #region GetMessageEncoding
    internal Encoding GetMessageEncoding()
    {
        var codepage = _reader.MessageCodepage;

        if (codepage is 0 or 1252) 
            return DefaultEncoding;

        try
        {
            return Encoding.GetEncoding(codepage);
        }
        catch
        {
            return DefaultEncoding;
        }
    }
    #endregion

    #region DecodeAnsiString
    private string DecodeAnsiString(byte[] bytes)
    {
        var length = bytes.Length;

        while (length > 0 && bytes[length - 1] == 0)
            length--;

        if (length == 0)
            return string.Empty;

        try
        {
            return GetMessageEncoding().GetString(bytes, 0, length);
        }
        catch
        {
            return DefaultEncoding.GetString(bytes, 0, length);
        }
    }
    #endregion

    #region ReadString
    private string ReadString()
    {
        var bytes = ReadByteArray();

        return DecodeAnsiString(bytes);
    }
    #endregion

    #region ReadAttrBytes
    private byte[] ReadAttrBytes()
    {
        return ReadBytes(RawValueLength);
    }
    #endregion

    #region ReadAttrString
    private string ReadAttrString()
    {
        var bytes = ReadBytes(RawValueLength);

        // attribute strings are null-terminated
        return DecodeAnsiString(bytes);
    }
    #endregion

    #region ReadAttrDateTime
    private DateTime ReadAttrDateTime()
    {
        int year = ReadInt16();
        int month = ReadInt16();
        int day = ReadInt16();
        int hour = ReadInt16();
        int minute = ReadInt16();
        int second = ReadInt16();
        // ReSharper disable once UnusedVariable
        int dow = ReadInt16();

        try
        {
            return new DateTime(year, month, day, hour, minute, second);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidDate, ex);
            return default;
        }
    }
    #endregion

    #region LoadPropertyName
    private void LoadPropertyName()
    {
        var guid = new Guid(ReadBytes(16));
        var kind = (NameIdKind)ReadInt32();

        if (kind == NameIdKind.Name)
        {
            var name = ReadUnicodeString();

            PropertyNameId = new NameId(guid, name);
        }
        else if (kind == NameIdKind.Id)
        {
            var id = ReadInt32();

            PropertyNameId = new NameId(guid, id);
        }
        else
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidAttributeValue);
            PropertyNameId = new NameId(guid, 0);
        }
    }
    #endregion

    #region ReadNextProperty
    /// <summary>
    ///     Advance to the next MAPI property.
    /// </summary>
    /// <remarks>
    ///     Advances to the next MAPI property.
    /// </remarks>
    /// <returns><c>true</c> if there is another property available to be read; otherwise <c>false</c>.</returns>
    /// <exception cref="MRTnefException">
    ///     The TNEF data is corrupt or invalid.
    /// </exception>
    public bool ReadNextProperty()
    {
        while (ReadNextValue())
        {
            // skip over the remaining value(s) for the current property...
        }

        if (_propertyIndex >= PropertyCount)
            return false;

        try
        {
            var type = (PropertyType)ReadInt16();
            var id = (PropertyId)ReadInt16();

            _propertyTag = new PropertyTag(id, type);

            if (_propertyTag.IsNamed)
                LoadPropertyName();

            LoadValueCount();
            _propertyIndex++;

            if (!TryGetPropertyValueLength(out _rawValueLength))
                return false;

            RawValueStreamOffset = _reader.StreamOffset;

            AttachMethod = id switch
            {
                PropertyId.AttachMethod => (AttachMethod)PeekInt32(),
                _ => AttachMethod
            };
        }
        catch (EndOfStreamException)
        {
            return false;
        }

        return CheckRawValueLength();
    }
    #endregion

    #region ReadNextRow
    /// <summary>
    ///     Advance to the next table row of properties.
    /// </summary>
    /// <remarks>
    ///     Advances to the next table row of properties.
    /// </remarks>
    /// <returns><c>true</c> if there is another row available to be read; otherwise <c>false</c>.</returns>
    /// <exception cref="MRTnefException">
    ///     The TNEF data is corrupt or invalid.
    /// </exception>
    public bool ReadNextRow()
    {
        while (ReadNextProperty())
        {
            // skip over the remaining property/properties in the current row...
        }

        if (_rowIndex >= RowCount)
            return false;

        try
        {
            LoadPropertyCount();
            _rowIndex++;
        }
        catch (EndOfStreamException)
        {
            _reader.SetComplianceError(ComplianceStatus.StreamTruncated);
            return false;
        }

        return true;
    }
    #endregion

    #region ReadNextValue
    /// <summary>
    ///     Advance to the next value in the TNEF stream.
    /// </summary>
    /// <remarks>
    ///     Advances to the next value in the TNEF stream.
    /// </remarks>
    /// <returns><c>true</c> if there is another value available to be read; otherwise <c>false</c>.</returns>
    /// <exception cref="MRTnefException">
    ///     The TNEF data is corrupt or invalid.
    /// </exception>
    public bool ReadNextValue()
    {
        if (_valueIndex >= ValueCount || PropertyCount == 0)
            return false;

        var offset = RawValueStreamOffset + RawValueLength;

        if (_reader.StreamOffset < offset && !_reader.Seek(offset))
            return false;

        try
        {
            if (!TryGetPropertyValueLength(out _rawValueLength))
                return false;

            RawValueStreamOffset = _reader.StreamOffset;
            _valueIndex++;
        }
        catch (EndOfStreamException)
        {
            return false;
        }

        return true;
    }
    #endregion

    #region ReadRawValue
    /// <summary>
    ///     Read the raw attribute or property value as a sequence of bytes.
    /// </summary>
    /// <remarks>
    ///     Reads the raw attribute or property value as a sequence of bytes.
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
    /// <exception cref="IOException">
    ///     An I/O error occurred.
    /// </exception>
    public int ReadRawValue(byte[] buffer, int offset, int count)
    {
        if (buffer is null)
            throw new ArgumentNullException(nameof(buffer));

        if (offset < 0 || offset >= buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(offset));

        if (count < 0 || count > buffer.Length - offset)
            throw new ArgumentOutOfRangeException(nameof(count));

        if (PropertyCount > 0 && _reader.StreamOffset == RawValueStreamOffset)
            switch (_propertyTag.ValueTnefType)
            {
                case PropertyType.Unicode:
                case PropertyType.String8:
                case PropertyType.Binary:
                case PropertyType.Object:
                    ReadInt32();
                    break;
            }

        var valueEndOffset = RawValueStreamOffset + RawValueLength;
        var valueLeft = valueEndOffset - _reader.StreamOffset;
        var n = Math.Min(valueLeft, count);

        return n > 0 ? _reader.ReadAttributeRawValue(buffer, offset, n) : 0;
    }
    #endregion

    #region ReadTextValue
    /// <summary>
    ///     Read the raw attribute or property value as a sequence of unicode characters.
    /// </summary>
    /// <remarks>
    ///     Reads the raw attribute or property value as a sequence of unicode characters.
    /// </remarks>
    /// <returns>
    ///     The total number of characters read into the buffer. This can be less than the number of characters
    ///     requested if that many bytes are not currently available, or zero (0) if the end of the stream has been
    ///     reached.
    /// </returns>
    /// <param name="buffer">The buffer to read data into.</param>
    /// <param name="offset">The offset into the buffer to start reading data.</param>
    /// <param name="count">The number of characters to read.</param>
    /// <exception cref="ArgumentNullException">
    ///     <paramref name="buffer" /> is <c>null</c>.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    ///     <para><paramref name="offset" /> is less than zero or greater than the length of <paramref name="buffer" />.</para>
    ///     <para>-or-</para>
    ///     <para>
    ///         The <paramref name="buffer" /> is not large enough to contain <paramref name="count" /> characters starting
    ///         at the specified <paramref name="offset" />.
    ///     </para>
    /// </exception>
    /// <exception cref="IOException">
    ///     An I/O error occurred.
    /// </exception>
    public int ReadTextValue(char[] buffer, int offset, int count)
    {
        if (buffer is null)
            throw new ArgumentNullException(nameof(buffer));

        if (offset < 0 || offset >= buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(offset));

        if (count < 0 || count > buffer.Length - offset)
            throw new ArgumentOutOfRangeException(nameof(count));

        if (_reader.StreamOffset == RawValueStreamOffset && _decoder is null)
            throw new InvalidOperationException();

        if (PropertyCount > 0 && _reader.StreamOffset == RawValueStreamOffset)
            switch (_propertyTag.ValueTnefType)
            {
                case PropertyType.Unicode:
                    ReadInt32();
                    _decoder = Encoding.Unicode.GetDecoder();
                    break;
                case PropertyType.String8:
                case PropertyType.Binary:
                case PropertyType.Object:
                    ReadInt32();
                    _decoder = GetMessageEncoding().GetDecoder();
                    break;
            }

        var valueEndOffset = RawValueStreamOffset + RawValueLength;
        var valueLeft = valueEndOffset - _reader.StreamOffset;
        var n = Math.Min(valueLeft, count);

        if (n <= 0)
            return 0;

        var bytes = new byte[n];

        n = _reader.ReadAttributeRawValue(bytes, 0, n);

        var flush = _reader.StreamOffset >= valueEndOffset;

        return _decoder.GetChars(bytes, 0, n, buffer, offset, flush);
    }
    #endregion

    #region TryGetPropertyValueLength
    private bool TryGetPropertyValueLength(out int length)
    {
        switch (_propertyTag.ValueTnefType)
        {
            case PropertyType.Unspecified:
            case PropertyType.Null:
                length = 0;
                break;
            case PropertyType.Boolean:
            case PropertyType.Error:
            case PropertyType.Long:
            case PropertyType.R4:
            case PropertyType.I2:
                length = 4;
                break;
            case PropertyType.Currency:
            case PropertyType.Double:
            case PropertyType.I8:
                length = 8;
                break;
            case PropertyType.ClassId:
                length = 16;
                break;
            case PropertyType.Unicode:
            case PropertyType.String8:
            case PropertyType.Binary:
            case PropertyType.Object:
                length = 4 + GetPaddedLength(PeekInt32());
                break;
            case PropertyType.AppTime:
            case PropertyType.SysTime:
                length = 8;
                break;
            default:
                _reader.SetComplianceError(ComplianceStatus.UnsupportedPropertyType);
                length = 0;

                return false;
        }

        return true;
    }
    #endregion

    #region GetPropertyValueType
    private Type GetPropertyValueType()
    {
        return _propertyTag.ValueTnefType switch
        {
            PropertyType.I2 => typeof(short),
            PropertyType.Boolean => typeof(bool),
            PropertyType.Currency => typeof(long),
            PropertyType.I8 => typeof(long),
            PropertyType.Error => typeof(int),
            PropertyType.Long => typeof(int),
            PropertyType.Double => typeof(double),
            PropertyType.R4 => typeof(float),
            PropertyType.AppTime => typeof(DateTime),
            PropertyType.SysTime => typeof(DateTime),
            PropertyType.Unicode => typeof(string),
            PropertyType.String8 => typeof(string),
            PropertyType.Binary => typeof(byte[]),
            PropertyType.ClassId => typeof(Guid),
            PropertyType.Object => typeof(byte[]),
            _ => typeof(object)
        };
    }
    #endregion

    #region GetAttributeValueType
    private Type GetAttributeValueType()
    {
        return _reader.AttributeType switch
        {
            AttributeType.Triples => typeof(byte[]),
            AttributeType.String => typeof(string),
            AttributeType.Text => typeof(string),
            AttributeType.Date => typeof(DateTime),
            AttributeType.Short => typeof(short),
            AttributeType.Long => typeof(int),
            AttributeType.Byte => typeof(byte[]),
            AttributeType.Word => typeof(short),
            AttributeType.DWord => typeof(int),
            _ => typeof(object)
        };
    }
    #endregion

    #region ReadPropertyValue
    private object ReadPropertyValue()
    {
        object value;

        switch (_propertyTag.ValueTnefType)
        {
            case PropertyType.Null:
                value = null;
                break;
            case PropertyType.I2:
                // 2 bytes for the short followed by 2 bytes of padding
                value = (short)(ReadInt32() & 0xFFFF);
                break;
            case PropertyType.Boolean:
                value = (ReadInt32() & 0xFF) != 0;
                break;
            case PropertyType.Currency:
            case PropertyType.I8:
                value = ReadInt64();
                break;
            case PropertyType.Error:
            case PropertyType.Long:
                value = ReadInt32();
                break;
            case PropertyType.Double:
                value = ReadDouble();
                break;
            case PropertyType.R4:
                value = ReadSingle();
                break;
            case PropertyType.AppTime:
                value = ReadAppTime();
                break;
            case PropertyType.SysTime:
                value = ReadSysTime();
                break;
            case PropertyType.Unicode:
                value = ReadUnicodeString();
                break;
            case PropertyType.String8:
                value = ReadString();
                break;
            case PropertyType.Binary:
                value = ReadByteArray();
                break;
            case PropertyType.ClassId:
                value = new Guid(ReadBytes(16));
                break;
            case PropertyType.Object:
                value = ReadByteArray();
                break;
            default:
                _reader.SetComplianceError(ComplianceStatus.UnsupportedPropertyType);
                value = null;
                break;
        }

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValue
    /// <summary>
    ///     Read the value.
    /// </summary>
    /// <remarks>
    ///     Reads an attribute or property value as its native type.
    /// </remarks>
    /// <returns>The value.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public object ReadValue()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        if (PropertyCount > 0)
            return ReadPropertyValue();

        object value = _reader.AttributeType switch
        {
            AttributeType.Triples => ReadAttrBytes(),
            AttributeType.String => ReadAttrString(),
            AttributeType.Text => ReadAttrString(),
            AttributeType.Date => ReadAttrDateTime(),
            AttributeType.Short => ReadInt16(),
            AttributeType.Long => ReadInt32(),
            AttributeType.Byte => ReadAttrBytes(),
            AttributeType.Word => ReadInt16(),
            AttributeType.DWord => ReadInt32(),
            _ => null
        };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsBoolean
    /// <summary>
    ///     Read the value as a boolean.
    /// </summary>
    /// <remarks>
    ///     Reads any integer-based attribute or property value as a boolean.
    /// </remarks>
    /// <returns>The value as a boolean.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a boolean.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public bool ReadValueAsBoolean()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        bool value;

        switch (PropertyCount)
        {
            case > 0:
                value = _propertyTag.ValueTnefType switch
                {
                    PropertyType.Boolean => (ReadInt32() & 0xFF) != 0,
                    PropertyType.I2 => (ReadInt32() & 0xFFFF) != 0,
                    PropertyType.Error => ReadInt32() != 0,
                    PropertyType.Long => ReadInt32() != 0,
                    PropertyType.Currency => ReadInt64() != 0,
                    PropertyType.I8 => ReadInt64() != 0,
                    _ => throw new InvalidOperationException()
                };

                break;
            default:
                value = _reader.AttributeType switch
                {
                    AttributeType.Short => ReadInt16() != 0,
                    AttributeType.Long => ReadInt32() != 0,
                    AttributeType.Word => ReadInt16() != 0,
                    AttributeType.DWord => ReadInt32() != 0,
                    AttributeType.Byte => ReadByte() != 0,
                    _ => throw new InvalidOperationException()
                };

                break;
        }

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsBytes
    /// <summary>
    ///     Read the value as a byte array.
    /// </summary>
    /// <remarks>
    ///     Reads any string, binary blob, Class ID, or Object attribute or property value as a byte array.
    /// </remarks>
    /// <returns>The value as a byte array.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a byte array.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public byte[] ReadValueAsBytes()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        byte[] bytes;

        if (PropertyCount > 0)
            bytes = _propertyTag.ValueTnefType switch
            {
                PropertyType.Unicode => ReadByteArray(),
                PropertyType.String8 => ReadByteArray(),
                PropertyType.Binary => ReadByteArray(),
                PropertyType.Object => ReadByteArray(),
                PropertyType.ClassId => ReadBytes(16),
                _ => throw new InvalidOperationException()
            };
        else
            bytes = _reader.AttributeType switch
            {
                AttributeType.Triples => ReadAttrBytes(),
                AttributeType.String => ReadAttrBytes(),
                AttributeType.Text => ReadAttrBytes(),
                AttributeType.Byte => ReadAttrBytes(),
                _ => throw new ArgumentOutOfRangeException()
            };

        _valueIndex++;

        return bytes;
    }
    #endregion

    #region ReadValueAsDateTime
    /// <summary>
    ///     Read the value as a date and time.
    /// </summary>
    /// <remarks>
    ///     Reads any date and time attribute or property value as a <see cref="DateTime" />.
    /// </remarks>
    /// <returns>The value as a date and time.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a date and time.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public DateTime ReadValueAsDateTime()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        DateTime value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.AppTime => ReadAppTime(),
                PropertyType.SysTime => ReadSysTime(),
                _ => throw new InvalidOperationException()
            };
        else if (_reader.AttributeType == AttributeType.Date)
            value = ReadAttrDateTime();
        else
            throw new InvalidOperationException();

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsDouble
    /// <summary>
    ///     Read the value as a double.
    /// </summary>
    /// <remarks>
    ///     Reads any numeric attribute or property value as a double.
    /// </remarks>
    /// <returns>The value as a double.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a double.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public double ReadValueAsDouble()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        double value;

        switch (PropertyCount)
        {
            case > 0:
                value = _propertyTag.ValueTnefType switch
                {
                    PropertyType.Boolean => ReadInt32() & 0xFF,
                    PropertyType.I2 => ReadInt32() & 0xFFFF,
                    PropertyType.Error => ReadInt32(),
                    PropertyType.Long => ReadInt32(),
                    PropertyType.Currency => ReadInt64(),
                    PropertyType.I8 => ReadInt64(),
                    PropertyType.Double => ReadDouble(),
                    PropertyType.R4 => ReadSingle(),
                    _ => throw new InvalidOperationException()
                };

                break;
            default:
                switch (_reader.AttributeType)
                {
                    case AttributeType.Short:
                        value = ReadInt16();
                        break;
                    case AttributeType.Long:
                        value = ReadInt32();
                        break;
                    case AttributeType.Word:
                        value = ReadInt16();
                        break;
                    case AttributeType.DWord:
                        value = ReadInt32();
                        break;
                    case AttributeType.Byte:
                        value = ReadDouble();
                        break;
                    default: throw new InvalidOperationException();
                }

                break;
        }

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsFloat
    /// <summary>
    ///     Read the value as a float.
    /// </summary>
    /// <remarks>
    ///     Reads any numeric attribute or property value as a float.
    /// </remarks>
    /// <returns>The value as a float.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a float.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public float ReadValueAsFloat()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        float value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.Boolean => ReadInt32() & 0xFF,
                PropertyType.I2 => ReadInt32() & 0xFFFF,
                PropertyType.Error => ReadInt32(),
                PropertyType.Long => ReadInt32(),
                PropertyType.Currency => ReadInt64(),
                PropertyType.I8 => ReadInt64(),
                PropertyType.Double => (float)ReadDouble(),
                PropertyType.R4 => ReadSingle(),
                _ => throw new InvalidOperationException()
            };
        else
            value = _reader.AttributeType switch
            {
                AttributeType.Short => ReadInt16(),
                AttributeType.Long => ReadInt32(),
                AttributeType.Word => ReadInt16(),
                AttributeType.DWord => ReadInt32(),
                AttributeType.Byte => ReadSingle(),
                _ => throw new InvalidOperationException()
            };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsGuid
    /// <summary>
    ///     Read the value as a GUID.
    /// </summary>
    /// <remarks>
    ///     Reads any Class ID value as a GUID.
    /// </remarks>
    /// <returns>The value as a GUID.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a GUID.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public Guid ReadValueAsGuid()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        Guid guid;

        if (PropertyCount > 0)
            guid = _propertyTag.ValueTnefType switch
            {
                PropertyType.ClassId => new Guid(ReadBytes(16)),
                _ => throw new InvalidOperationException()
            };
        else
            throw new InvalidOperationException();

        _valueIndex++;

        return guid;
    }
    #endregion

    #region ReadValueAsInt16
    /// <summary>
    ///     Read the value as a 16-bit integer.
    /// </summary>
    /// <remarks>
    ///     Reads any integer-based attribute or property value as a 16-bit integer.
    /// </remarks>
    /// <returns>The value as a 16-bit integer.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a 16-bit integer.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public short ReadValueAsInt16()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        short value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.Boolean => (short)(ReadInt32() & 0xFF),
                PropertyType.I2 => (short)(ReadInt32() & 0xFFFF),
                PropertyType.Error => (short)ReadInt32(),
                PropertyType.Long => (short)ReadInt32(),
                PropertyType.Currency => (short)ReadInt64(),
                PropertyType.I8 => (short)ReadInt64(),
                PropertyType.Double => (short)ReadDouble(),
                PropertyType.R4 => (short)ReadSingle(),
                _ => throw new InvalidOperationException()
            };
        else
            value = _reader.AttributeType switch
            {
                AttributeType.Short => ReadInt16(),
                AttributeType.Long => (short)ReadInt32(),
                AttributeType.Word => ReadInt16(),
                AttributeType.DWord => (short)ReadInt32(),
                AttributeType.Byte => ReadInt16(),
                _ => throw new InvalidOperationException()
            };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsInt32
    /// <summary>
    ///     Read the value as a 32-bit integer.
    /// </summary>
    /// <remarks>
    ///     Reads any integer-based attribute or property value as a 32-bit integer.
    /// </remarks>
    /// <returns>The value as a 32-bit integer.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a 32-bit integer.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public int ReadValueAsInt32()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        int value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.Boolean => ReadInt32() & 0xFF,
                PropertyType.I2 => ReadInt32() & 0xFFFF,
                PropertyType.Error => ReadInt32(),
                PropertyType.Long => ReadInt32(),
                PropertyType.Currency => (int)ReadInt64(),
                PropertyType.I8 => (int)ReadInt64(),
                PropertyType.Double => (int)ReadDouble(),
                PropertyType.R4 => (int)ReadSingle(),
                _ => throw new InvalidOperationException()
            };
        else
            value = _reader.AttributeType switch
            {
                AttributeType.Short => ReadInt16(),
                AttributeType.Long => ReadInt32(),
                AttributeType.Word => ReadInt16(),
                AttributeType.DWord => ReadInt32(),
                AttributeType.Byte => ReadInt32(),
                _ => throw new InvalidOperationException()
            };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsInt64
    /// <summary>
    ///     Read the value as a 64-bit integer.
    /// </summary>
    /// <remarks>
    ///     Reads any integer-based attribute or property value as a 64-bit integer.
    /// </remarks>
    /// <returns>The value as a 64-bit integer.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a 64-bit integer.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public long ReadValueAsInt64()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        long value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.Boolean => ReadInt32() & 0xFF,
                PropertyType.I2 => ReadInt32() & 0xFFFF,
                PropertyType.Error => ReadInt32(),
                PropertyType.Long => ReadInt32(),
                PropertyType.Currency => ReadInt64(),
                PropertyType.I8 => ReadInt64(),
                PropertyType.Double => (long)ReadDouble(),
                PropertyType.R4 => (long)ReadSingle(),
                _ => throw new InvalidOperationException()
            };
        else
            value = _reader.AttributeType switch
            {
                AttributeType.Short => ReadInt16(),
                AttributeType.Long => ReadInt32(),
                AttributeType.Word => ReadInt16(),
                AttributeType.DWord => ReadInt32(),
                AttributeType.Byte => ReadInt64(),
                _ => throw new InvalidOperationException()
            };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsString
    /// <summary>
    ///     Read the value as a string.
    /// </summary>
    /// <remarks>
    ///     Reads any string or binary blob values as a string.
    /// </remarks>
    /// <returns>The value as a string.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a string.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    public string ReadValueAsString()
    {
        if (_valueIndex >= ValueCount || _reader.StreamOffset > RawValueStreamOffset)
            throw new InvalidOperationException();

        string value;

        if (PropertyCount > 0)
            value = _propertyTag.ValueTnefType switch
            {
                PropertyType.Unicode => ReadUnicodeString(),
                PropertyType.String8 => ReadString(),
                PropertyType.Binary => ReadString(),
                _ => throw new InvalidOperationException()
            };
        else
            value = _reader.AttributeType switch
            {
                AttributeType.String => ReadAttrString(),
                AttributeType.Text => ReadAttrString(),
                AttributeType.Byte => ReadAttrString(),
                _ => throw new InvalidOperationException()
            };

        _valueIndex++;

        return value;
    }
    #endregion

    #region ReadValueAsUri
    /// <summary>
    ///     Read the value as a Uri.
    /// </summary>
    /// <remarks>
    ///     Reads any string or binary blob values as a Uri.
    /// </remarks>
    /// <returns>The value as a Uri.</returns>
    /// <exception cref="InvalidOperationException">
    ///     There are no more values to read or the value could not be read as a string.
    /// </exception>
    /// <exception cref="EndOfStreamException">
    ///     The TNEF stream is truncated and the value could not be read.
    /// </exception>
    internal Uri ReadValueAsUri()
    {
        var value = ReadValueAsString();

        if (Uri.IsWellFormedUriString(value, UriKind.Absolute))
            return new Uri(value, UriKind.Absolute);

        return Uri.IsWellFormedUriString(value, UriKind.Relative) ? new Uri(value, UriKind.Relative) : null;
    }
    #endregion

    #region GetHashCode
    /// <summary>
    ///     Serves as a hash function for a <see cref="PropertyReader" /> object.
    /// </summary>
    /// <remarks>
    ///     Serves as a hash function for a <see cref="PropertyReader" /> object.
    /// </remarks>
    /// <returns>
    ///     A hash code for this instance that is suitable for use in hashing algorithms
    ///     and data structures such as a hash table.
    /// </returns>
    public override int GetHashCode()
    {
        return _reader.GetHashCode();
    }
    #endregion

    #region Equals
    /// <summary>
    ///     Determine whether the specified <see cref="object" /> is equal to the current <see cref="PropertyReader" />.
    /// </summary>
    /// <remarks>
    ///     Determines whether the specified <see cref="object" /> is equal to the current <see cref="PropertyReader" />.
    /// </remarks>
    /// <param name="obj">The <see cref="object" /> to compare with the current <see cref="PropertyReader" />.</param>
    /// <returns>
    ///     <c>true</c> if the specified <see cref="object" /> is equal to the current
    ///     <see cref="PropertyReader" />; otherwise, <c>false</c>.
    /// </returns>
    public override bool Equals(object obj)
    {
        return obj is PropertyReader prop && prop._reader == _reader;
    }
    #endregion

    #region MyReLoadPropertyCountgion
    private void LoadPropertyCount()
    {
        if ((PropertyCount = ReadInt32()) < 0)
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidPropertyLength);
            PropertyCount = 0;
        }

        _propertyIndex = 0;
        ValueCount = 0;
        _valueIndex = 0;
        _decoder = null;
    }
    #endregion

    #region ReadValueCount
    private int ReadValueCount()
    {
        int count;

        if ((count = ReadInt32()) < 0)
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidAttributeValue);
            return 0;
        }

        return count;
    }
    #endregion

    #region LoadValueCount
    private void LoadValueCount()
    {
        if (_propertyTag.IsMultiValued)
            ValueCount = ReadValueCount();
        else
            switch (_propertyTag.ValueTnefType)
            {
                case PropertyType.Unicode:
                case PropertyType.String8:
                case PropertyType.Binary:
                case PropertyType.Object:
                    ValueCount = ReadValueCount();
                    break;
                default:
                    ValueCount = 1;
                    break;
            }

        _valueIndex = 0;
        _decoder = null;
    }
    #endregion

    #region LoadRowCount
    private void LoadRowCount()
    {
        if ((RowCount = ReadInt32()) < 0)
        {
            _reader.SetComplianceError(ComplianceStatus.InvalidRowCount);
            RowCount = 0;
        }

        PropertyCount = 0;
        _propertyIndex = 0;
        ValueCount = 0;
        _valueIndex = 0;
        _decoder = null;
        _rowIndex = 0;
    }
    #endregion

    #region Load
    internal void Load()
    {
        _propertyTag = PropertyTag.Null;
        RawValueStreamOffset = 0;
        _rawValueLength = 0;
        PropertyCount = 0;
        _propertyIndex = 0;
        ValueCount = 0;
        _valueIndex = 0;
        _decoder = null;
        RowCount = 0;
        _rowIndex = 0;

        switch (_reader.AttributeTag)
        {
            case AttributeTag.MapiProperties:
            case AttributeTag.Attachment:
                LoadPropertyCount();
                break;
            case AttributeTag.RecipientTable:
                LoadRowCount();
                break;
            default:
                _rawValueLength = _reader.AttributeRawValueLength;
                RawValueStreamOffset = _reader.StreamOffset;
                ValueCount = 1;
                break;
        }
    }
    #endregion
}