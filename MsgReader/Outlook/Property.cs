﻿//
// Property.cs
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
using System.IO;
using System.Text;
using MsgReader.Exceptions;
using MsgReader.Helpers;

// ReSharper disable UnusedMember.Global

// ReSharper disable InconsistentNaming

namespace MsgReader.Outlook;

/// <summary>
///     Pointer to a variable of type SPropValue that specifies the property value array describing the properties
///     for the recipient. The rgPropVals member can be NULL.
/// </summary>
public class Property
{
    #region Properties
    /// <summary>
    ///     The id of the property
    /// </summary>
    internal ushort Id { get; }

    /// <summary>
    ///     Returns the Property as a readable string without the stream prefix and type
    /// </summary>
    /// <returns></returns>
    public string ShortName => Id.ToString("X4");

    /// <summary>
    ///     Returns the Property as a readable string
    /// </summary>
    /// <returns></returns>
    public string Name => MapiTags.SubStorageStreamPrefix + Id.ToString("X4") + ((ushort)Type).ToString("X4");

    /// <summary>
    ///     The <see cref="MsgReader.Outlook.PropertyType" />
    /// </summary>
    internal PropertyType Type { get; }

    /// <summary>
    ///     Returns <c>true</c> when this property is part of a multi value property
    /// </summary>
    internal bool MultiValue { get; }

    /// <summary>
    ///     The property data
    /// </summary>
    internal byte[] Data { get; }

    /// <summary>
    ///     Returns <see cref="Data" /> as an integer when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_SHORT" />,
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONG" /> or <see cref="MsgReader.Outlook.PropertyType.PT_ERROR" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not <see cref="MsgReader.Outlook.PropertyType.PT_SHORT" /> or
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONG" />
    /// </exception>
    internal int ToInt
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_SHORT => BitConverter.ToInt16(Data, 0),
                PropertyType.PT_LONG => BitConverter.ToInt32(Data, 0),
                _ => throw new MRInvalidProperty("Type is not PT_SHORT or PT_LONG")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
    /// </exception>
    internal float ToSingle
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_FLOAT => BitConverter.ToSingle(Data, 0),
                _ => throw new MRInvalidProperty("Type is not PT_FLOAT")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_DOUBLE" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_DOUBLE" />
    /// </exception>
    internal double ToDouble
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_DOUBLE => BitConverter.ToDouble(Data, 0),
                _ => throw new MRInvalidProperty("Type is not PT_DOUBLE")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a decimal when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
    /// </exception>
    internal decimal ToDecimal
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_FLOAT => ByteArrayToDecimal(Data, 0),
                _ => throw new MRInvalidProperty("Type is not PT_FLOAT")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a <see cref="DateTime"/>> when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_APPTIME" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_APPTIME" /> or
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_SYSTIME" />
    /// </exception>
    internal DateTimeOffset ToDateTime
    {
        get
        {
            switch (Type)
            {
                case PropertyType.PT_APPTIME:
                    var oaDate = BitConverter.ToDouble(Data, 0);
                    return DateTime.FromOADate(oaDate);

                default:
                    throw new MRInvalidProperty("Type is not a PT_APPTIME");
            }
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a <see cref="DateTimeOffset"/> when <see cref="Type" /> is set to
    ///     or <see cref="MsgReader.Outlook.PropertyType.PT_SYSTIME" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_APPTIME" /> or
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_SYSTIME" />
    /// </exception>
    internal DateTimeOffset ToDateTimeOffset
    {
        get
        {
            switch (Type)
            {
                case PropertyType.PT_SYSTIME:
                    var fileTime = BitConverter.ToInt64(Data, 0);
                    return DateTimeOffset.FromFileTime(fileTime);

                default:
                    throw new MRInvalidProperty("Type is not a PT_SYSTIME");
            }
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_BOOLEAN" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_BOOLEAN" />
    /// </exception>
    internal bool ToBool
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_BOOLEAN => BitConverter.ToBoolean(Data, 0),
                _ => throw new MRInvalidProperty("Type is not PT_BOOLEAN")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONGLONG" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONGLONG" />
    /// </exception>
    internal long ToLong
    {
        get
        {
            switch (Type)
            {
                case PropertyType.PT_LONG:
                case PropertyType.PT_LONGLONG:
                    return BitConverter.ToInt64(Data, 0);

                default:
                    throw new MRInvalidProperty("Type is not PtypInteger64");
            }
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a string when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_UNICODE" />
    ///     or <see cref="MsgReader.Outlook.PropertyType.PT_STRING8" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_UNICODE" />
    ///     or <see cref="MsgReader.Outlook.PropertyType.PT_STRING8" />
    /// </exception>
    public new string ToString
    {
        get
        {
            switch (Type)
            {
                case PropertyType.PT_UNICODE:
                case PropertyType.PT_STRING8:
                    var encoding = Type == PropertyType.PT_STRING8 ? Encoding.Default : Encoding.Unicode;
                    using (var recyclableMemoryStream = StreamHelpers.Manager.GetStream("Property.cs", Data, 0, Data.Length))
                    using (var streamReader = new StreamReader(recyclableMemoryStream, encoding))
                    {
                        var streamContent = streamReader.ReadToEnd();
                        return streamContent.TrimEnd('\0');
                    }

                default:
                    var encoding2 = Type == PropertyType.PT_STRING8 ? Encoding.Default : Encoding.Unicode;
                    using (var recyclableMemoryStream = StreamHelpers.Manager.GetStream("Property.cs", Data, 0, Data.Length))
                    using (var streamReader = new StreamReader(recyclableMemoryStream, encoding2))
                    {
                        var streamContent = streamReader.ReadToEnd();
                        return streamContent.TrimEnd('\0');
                    }
                //throw new MRInvalidProperty("Type is not PT_UNICODE or PT_STRING8");
            }
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a byte[] when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_BINARY" />
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_OBJECT" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_BINARY" />
    /// </exception>
    public byte[] ToBinary
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_BINARY => Data,
                _ => throw new MRInvalidProperty("Type is not PT_BINARY")
            };
        }
    }

    /// <summary>
    ///     Returns <see cref="Data" /> as a Guid when <see cref="Type" /> is set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_CLSID" />
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_OBJECT" />
    /// </summary>
    /// <exception cref="MRInvalidProperty">
    ///     Raised when the <see cref="Type" /> is not set to
    ///     <see cref="MsgReader.Outlook.PropertyType.PT_BINARY" />
    /// </exception>
    public Guid ToGuid
    {
        get
        {
            return Type switch
            {
                PropertyType.PT_CLSID => new Guid(Data),
                _ => throw new MRInvalidProperty("Type is not PT_CLSID")
            };
        }
    }
    #endregion

    #region ByteArrayToDecimal
    /// <summary>
    ///     Converts a byte array to a decimal
    /// </summary>
    /// <param name="source">The byte array</param>
    /// <param name="offset">The offset to start reading</param>
    /// <returns></returns>
    private static decimal ByteArrayToDecimal(byte[] source, int offset)
    {
        var i1 = BitConverter.ToInt32(source, offset);
        var i2 = BitConverter.ToInt32(source, offset + 4);
        var i3 = BitConverter.ToInt32(source, offset + 8);
        var i4 = BitConverter.ToInt32(source, offset + 12);

        return new decimal([i1, i2, i3, i4]);
    }
    #endregion

    #region Constructor
    /// <summary>
    ///     Creates this object and sets all its properties
    /// </summary>
    /// <param name="id">The id of the property</param>
    /// <param name="type">The <see cref="MsgReader.Outlook.PropertyType" /></param>
    /// <param name="data">The property data</param>
    /// <param name="multiValue">
    ///     Set to <c>true</c> to indicate that this property is part of a
    ///     multi value property
    /// </param>
    public Property(ushort id, PropertyType type, byte[] data, bool multiValue = false)
    {
        Id = id;
        Type = type;
        Data = data;
        MultiValue = multiValue;
    }
    #endregion
}