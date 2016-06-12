using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using MsgReader.Exceptions;

// ReSharper disable InconsistentNaming

/*
   Copyright 2015 - 2016 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
{
    #region PropertyFlags
    /// <summary>
    ///     Flags used to set on a <see cref="Property" />
    /// </summary>
    [Flags]
    internal enum PropertyFlag : uint
    {
        // ReSharper disable InconsistentNaming
        /// <summary>
        ///     If this flag is set for a property, that property MUST NOT be deleted from the .msg file
        ///     (irrespective of which storage it is contained in) and implementations MUST return an error
        ///     if any attempt is made to do so. This flag is set in circumstances where the implementation
        ///     depends on that property always being present in the .msg file once it is written there.
        /// </summary>
        PROPATTR_MANDATORY = 0x00000001,

        /// <summary>
        ///     If this flag is not set on a property, that property MUST NOT be read from the .msg file
        ///     and implementations MUST return an error if any attempt is made to read it. This flag is
        ///     set on all properties unless there is an implementation-specific reason to prevent a property
        ///     from being read from the .msg file.
        /// </summary>
        PROPATTR_READABLE = 0x00000002,

        /// <summary>
        ///     If this flag is not set on a property, that property MUST NOT be modified or deleted and
        ///     implementations MUST return an error if any attempt is made to do so. This flag is set in
        ///     circumstances where the implementation depends on the properties being writable.
        /// </summary>
        PROPATTR_WRITABLE = 0x00000004
        // ReSharper restore InconsistentNaming
    }
    #endregion

    #region PropertyType
    /// <summary>
    ///     The type of a property in the properties stream
    /// </summary>
    internal enum PropertyType : ushort
    {
        /// <summary>
        ///     2 bytes; a 16-bit integer (PT_I2, i2, ui2)
        /// </summary>
        PT_SHORT = 0x0002,

        /// <summary>
        ///     4 bytes; a 32-bit integer (PT_LONG, PT_I4, int, ui4)
        /// </summary>
        PT_LONG = 0x0003,

        /// <summary>
        ///     4 bytes; a 32-bit floating point number (PT_FLOAT, PT_R4, float, r4)
        /// </summary>
        PT_FLOAT = 0x0004,

        /// <summary>
        ///     8 bytes; a 64-bit floating point number (PT_DOUBLE, PT_R8, r8)
        /// </summary>
        PT_DOUBLE = 0x0005,

        /// <summary>
        ///     8 bytes; a 64-bit floating point number in which the whole number part represents the number of days since
        ///     December 30, 1899, and the fractional part represents the fraction of a day since midnight (PT_APPTIME)
        /// </summary>
        PT_APPTIME = 0x0007,

        /// <summary>
        ///     4 bytes; a 32-bit integer encoding error information as specified in section 2.4.1. (PT_ERROR)
        /// </summary>
        PT_ERROR = 0x000A,

        /// <summary>
        ///     1 byte; restricted to 1 or 0 (PT_BOOLEAN. bool)
        /// </summary>
        PT_BOOLEAN = 0x000B,

        /// <summary>
        ///     8 bytes; a 64-bit integer (PT_LONGLONG, PT_I8, i8, ui8)
        /// </summary>
        PT_I8 = 0x0014,

        /// <summary>
        ///     8 bytes; a 64-bit integer (PT_LONGLONG, PT_I8, i8, ui8)
        /// </summary>
        PT_LONGLONG = 0x0014,

        /// <summary>
        ///     Variable size; a string of Unicode characters in UTF-16LE format encoding with terminating null character
        ///     (0x0000). (PT_UNICODE, string)
        /// </summary>
        PT_UNICODE = 0x001F,

        /// <summary>
        ///     Variable size; a string of multibyte characters in externally specified encoding with terminating null
        ///     character (single 0 byte). (PT_STRING8) ... ANSI format
        /// </summary>
        PT_STRING8 = 0x001E,

        /// <summary>
        ///     8 bytes; a 64-bit integer representing the number of 100-nanosecond intervals since January 1, 1601
        ///     (PT_SYSTIME, time, datetime, datetime.tz, datetime.rfc1123, Date, time, time.tz)
        /// </summary>
        PT_SYSTIME = 0x0040,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many bytes. (PT_BINARY)
        /// </summary>
        PT_BINARY = 0x0102,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_SHORT values. (PT_MV_SHORT, PT_MV_I2, mv.i2)
        /// </summary>
        PT_MV_SHORT = 0x1002,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_LONG values. (PT_MV_LONG, PT_MV_I4, mv.i4)
        /// </summary>
        PT_MV_LONG = 0x1003,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_FLOAT values. (PT_MV_FLOAT, PT_MV_R4, mv.float)
        /// </summary>
        PT_MV_FLOAT = 0x1004,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_DOUBLE values. (PT_MV_DOUBLE, PT_MV_R8)
        /// </summary>
        PT_MV_DOUBLE = 0x1005,

        ///// <summary>
        /////     Variable size; a COUNT field followed by that many PT_MV_CURRENCY values. (PT_MV_CURRENCY, mv.fixed.14.4)
        ///// </summary>
        //PT_MV_CURRENCY = 0x1006,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_APPTIME values. (PT_MV_APPTIME)
        /// </summary>
        PT_MV_APPTIME = 0x1007,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_I8 values. (PT_MV_I8, PT_MV_LONGLONG)
        /// </summary>
        PT_MV_I8 = 0x1014,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_UNICODE values. (PT_MV_UNICODE)
        /// </summary>
        PT_MV_TSTRING = 0x101F,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_UNICODE values. (PT_MV_UNICODE)
        /// </summary>
        PT_MV_UNICODE = 0x101F,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_STRING8 values. (PT_MV_STRING8, mv.string)
        /// </summary>
        PT_MV_STRING8 = 0x101E,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_SYSTIME values. (PT_MV_SYSTIME)
        /// </summary>
        PT_MV_SYSTIME = 0x1040,

        ///// <summary>
        /////     Variable size; a COUNT field followed by that many PT_MV_CLSID values. (PT_MV_CLSID, mv.uuid)
        ///// </summary>
        //PT_MV_CLSID = 0x1048,

        /// <summary>
        ///     Variable size; a COUNT field followed by that many PT_MV_BINARY values. (PT_MV_BINARY, mv.bin.hex)
        /// </summary>
        PT_MV_BINARY = 0x1102,

        /// <summary>
        ///     Any: this property type value matches any type; a server MUST return the actual type in its response. Servers
        ///     MUST NOT return this type in response to a client request other than NspiGetIDsFromNames or the
        ///     RopGetPropertyIdsFromNamesROP request ([MS-OXCROPS] section 2.2.8.1). (PT_UNSPECIFIED)
        /// </summary>
        PT_UNSPECIFIED = 0x0000,

        /// <summary>
        ///     None: This property is a placeholder. (PT_NULL)
        /// </summary>
        PT_NULL = 0x0001,

        /// <summary>
        ///     The property value is a Component Object Model (COM) object, as specified in section 2.11.1.5. (PT_OBJECT)
        /// </summary>
        PT_OBJECT = 0x000D
    }
    #endregion

    /// <summary>
    ///     A property inside the MSG file
    /// </summary>
    internal class Property
    {
        #region Properties
        /// <summary>
        ///     The id of the property
        /// </summary>
        internal ushort Id { get; private set; }

        /// <summary>
        ///     Returns the Property as a readable string
        /// </summary>
        /// <returns></returns>
        public string Name
        {
            get { return MapiTags.SubStorageStreamPrefix + Id.ToString("X4") + ((ushort)Type).ToString("X4"); }
        }

        /// <summary>
        ///     The <see cref="PropertyType" />
        /// </summary>
        internal PropertyType Type { get; private set; }

        /// <summary>
        ///     The <see cref="PropertyFlag">property flags</see> that have been set
        ///     in its <see cref="uint" /> raw form
        /// </summary>
        internal uint Flags { get; private set; }

        /// <summary>
        ///     The <see cref="PropertyFlag">property flags</see> that have been set
        ///     as a readonly collection
        /// </summary>
        internal ReadOnlyCollection<PropertyFlag> FlagsCollection
        {
            get
            {
                var result = new List<PropertyFlag>();

                if ((Flags & Convert.ToUInt32(PropertyFlag.PROPATTR_MANDATORY)) != 0)
                    result.Add(PropertyFlag.PROPATTR_MANDATORY);

                if ((Flags & Convert.ToUInt32(PropertyFlag.PROPATTR_READABLE)) != 0)
                    result.Add(PropertyFlag.PROPATTR_READABLE);

                if ((Flags & Convert.ToUInt32(PropertyFlag.PROPATTR_WRITABLE)) != 0)
                    result.Add(PropertyFlag.PROPATTR_WRITABLE);

                return result.AsReadOnly();
            }
        }

        /// <summary>
        ///     The property data
        /// </summary>
        internal byte[] Data { get; private set; }

        /// <summary>
        ///     Returns <see cref="Data" /> as an integer when <see cref="Type" /> is set to
        ///     <see cref="PropertyType.PT_SHORT" />,
        ///     <see cref="PropertyType.PT_LONG" /> or <see cref="PropertyType.PT_ERROR" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="PropertyType.PT_SHORT"/> or
        /// <see cref="PropertyType.PT_LONG"/></exception>
        internal int ToInt
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_SHORT:
                        return BitConverter.ToInt16(Data, 0);

                    case PropertyType.PT_LONG:
                        return BitConverter.ToInt32(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypInteger16 or PtypInteger32");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to 
        ///     <see cref="PropertyType.PT_FLOAT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="PropertyType.PT_FLOAT"/></exception>
        internal float ToSingle
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_FLOAT:
                        return BitConverter.ToSingle(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypFloating32");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to 
        ///     <see cref="PropertyType.PT_DOUBLE" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="PropertyType.PT_DOUBLE"/></exception>
        internal double ToDouble
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_DOUBLE:
                        return BitConverter.ToDouble(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypFloating64");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a decimal when <see cref="Type" /> is set to
        ///     <see cref="PropertyType.PT_FLOAT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="PropertyType.PT_FLOAT"/></exception>
        internal decimal ToDecimal
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_FLOAT:
                        return ByteArrayToDecimal(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypFloating32");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a datetime when <see cref="Type" /> is set to
        ///     <see cref="PropertyType.PT_APPTIME" />
        ///     or <see cref="PropertyType.PT_SYSTIME" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="PropertyType.PT_APPTIME"/> or
        /// <see cref="PropertyType.PT_SYSTIME"/></exception>
        internal DateTime ToDateTime
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_APPTIME:
                        var oaDate = BitConverter.ToDouble(Data, 0);
                        return DateTime.FromOADate(oaDate);
                        
                    case PropertyType.PT_SYSTIME:
                        var fileTime = BitConverter.ToInt64(Data, 0);
                        return DateTime.FromFileTime(fileTime);

                    default:
                        throw new MRInvalidProperty("Type is not PtypFloating32");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to <see cref="PropertyType.PT_BOOLEAN" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="PropertyType.PT_BOOLEAN"/></exception>
        internal bool ToBool
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_BOOLEAN:
                        return BitConverter.ToBoolean(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypBoolean");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to
        ///     <see cref="PropertyType.PT_LONGLONG" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="PropertyType.PT_LONGLONG"/></exception>
        internal long ToLong
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_LONGLONG:
                        return BitConverter.ToInt64(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PtypInteger64");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a string when <see cref="Type" /> is set to <see cref="PropertyType.PT_UNICODE" />
        ///     or <see cref="PropertyType.PT_STRING8" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="PropertyType.PT_UNICODE"/> or <see cref="PropertyType.PT_STRING8" /></exception>
        public new string ToString
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_UNICODE:
                    case PropertyType.PT_STRING8:
                        var encoding = Type == PropertyType.PT_STRING8 ? Encoding.Default : Encoding.Unicode;
                        using (var memoryStream = new MemoryStream(Data))
                        using (var streamReader = new StreamReader(memoryStream, encoding))
                        {
                            var streamContent = streamReader.ReadToEnd();
                            return streamContent.TrimEnd('\0');
                        }

                    default:
                        throw new MRInvalidProperty("Type is not PtypString or PtypString8");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a byte[] when <see cref="Type" /> is set to <see cref="PropertyType.PT_BINARY" />
        ///     <see cref="PropertyType.PT_OBJECT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="PropertyType.PT_BINARY"/></exception>
        public byte[] ToBinary
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_BINARY:
                        return Data;

                    default:
                        throw new MRInvalidProperty("Type is not PtypBinary");
                }
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

            return new decimal(new[] { i1, i2, i3, i4 });
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all its propertues
        /// </summary>
        /// <param name="id">The id of the property</param>
        /// <param name="type">The <see cref="PropertyType" /></param>
        /// <param name="flags">The <see cref="PropertyFlag" /></param>
        /// <param name="data">The property data</param>
        internal Property(ushort id, PropertyType type, PropertyFlag flags, byte[] data)
        {
            Id = id;
            Type = type;
            Flags = Convert.ToUInt32(flags);
            Data = data;
        }

        /// <summary>
        ///     Creates this object and sets all its propertues
        /// </summary>
        /// <param name="id">The id of the property</param>
        /// <param name="type">The <see cref="PropertyType" /></param>
        /// <param name="flags">The <see cref="PropertyFlag" /></param>
        /// <param name="data">The property data</param>
        internal Property(ushort id, PropertyType type, uint flags, byte[] data)
        {
            Id = id;
            Type = type;
            Flags = flags;
            Data = data;
        }
        #endregion
    }
}