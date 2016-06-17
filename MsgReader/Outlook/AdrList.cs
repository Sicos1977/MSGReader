using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using MsgReader.Exceptions;
using MsgReader.Helpers;
// ReSharper disable InconsistentNaming

namespace MsgReader.Outlook
{
    #region Class AdrList
    /// <summary>
    ///     Describes zero or more properties belonging to one or more recipients. 
    /// </summary>
    /// <remarks>
    /// https://msdn.microsoft.com/en-us/library/ms629442%28VS.85%29.aspx
    /// </remarks>
    public class AdrList : List<SProperty>
    {
        #region Constructor
        internal AdrList(byte[] data)
        {
            var binaryReader = new BinaryReader(new MemoryStream(data));
            ReadProperties(binaryReader);
        }
        #endregion

        #region ReadProperties
        /// <summary>
        ///     Reads all the <see cref="SProperty" /> objects from the given <paramref name="binaryReader" />
        /// </summary>
        /// <param name="binaryReader"></param>
        internal void ReadProperties(BinaryReader binaryReader)
        {
            // RowCount (4 bytes):  An integer that specifies the number of structures 
            // in the RecipientRow field.
            var rows = binaryReader.ReadUInt32();

            // Count amount of entries in the adrlist
            var entries = binaryReader.ReadUInt64();

            while (!binaryReader.Eos())
            {
                // property tag: A 32-bit value that contains a property type and a property ID. The low-order 16 bits 
                // represent the property type. The high-order 16 bits represent the property ID.
                var type = (PropertyType)binaryReader.ReadUInt16();
                var id = binaryReader.ReadUInt16();
                var length = binaryReader.ReadUInt16();
                var data = binaryReader.ReadBytes(length);

                var pos = binaryReader.BaseStream.Position.ToString("X4");
                // The dwAlignPad member is used as padding to make sure proper alignment on computers that require 8-byte alignment
                // for 8-byte values. Developers who write code on such computers should use memory allocation routines that allocate 
                // the SPropValue arrays on 8-byte boundaries.
                while (binaryReader.PeekChar() == 0)
                    binaryReader.ReadByte();

                var prop = new SProperty(id, type, data);
#if (DEBUG)
                if (prop.Type == PropertyType.PT_UNICODE || prop.Type == PropertyType.PT_STRING8)
                    Debug.WriteLine(string.Format("{0} - {1}", prop.Name, prop.ToString));
#endif
                Add(prop);
            }
        }
        #endregion
    }
    #endregion

    /// <summary>
    ///     Pointer to a variable of type SPropValue that specifies the property value array describing the properties 
    ///     for the recipient. The rgPropVals member can be NULL.
    /// </summary>
    public class SProperty
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
        ///     The <see cref="MsgReader.Outlook.PropertyType" />
        /// </summary>
        internal PropertyType Type { get; private set; }

        /// <summary>
        ///     The property data
        /// </summary>
        internal byte[] Data { get; private set; }

        /// <summary>
        ///     Returns <see cref="Data" /> as an integer when <see cref="Type" /> is set to
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_SHORT" />,
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONG" /> or <see cref="MsgReader.Outlook.PropertyType.PT_ERROR" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="MsgReader.Outlook.PropertyType.PT_SHORT"/> or
        /// <see cref="MsgReader.Outlook.PropertyType.PT_LONG"/></exception>
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
                        throw new MRInvalidProperty("Type is not PT_SHORT or PT_LONG");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to 
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT"/></exception>
        internal float ToSingle
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_FLOAT:
                        return BitConverter.ToSingle(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PT_FLOAT");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a single when <see cref="Type" /> is set to 
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_DOUBLE" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="MsgReader.Outlook.PropertyType.PT_DOUBLE"/></exception>
        internal double ToDouble
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_DOUBLE:
                        return BitConverter.ToDouble(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PT_DOUBLE");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a decimal when <see cref="Type" /> is set to
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not <see cref="MsgReader.Outlook.PropertyType.PT_FLOAT"/></exception>
        internal decimal ToDecimal
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_FLOAT:
                        return ByteArrayToDecimal(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PT_FLOAT");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a datetime when <see cref="Type" /> is set to
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_APPTIME" />
        ///     or <see cref="MsgReader.Outlook.PropertyType.PT_SYSTIME" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_APPTIME"/> or
        /// <see cref="MsgReader.Outlook.PropertyType.PT_SYSTIME"/></exception>
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
                        throw new MRInvalidProperty("Type is not PT_APPTIME or PT_SYSTIME");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to <see cref="MsgReader.Outlook.PropertyType.PT_BOOLEAN" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_BOOLEAN"/></exception>
        internal bool ToBool
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_BOOLEAN:
                        return BitConverter.ToBoolean(Data, 0);

                    default:
                        throw new MRInvalidProperty("Type is not PT_BOOLEAN");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a boolean when <see cref="Type" /> is set to
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_LONGLONG" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_LONGLONG"/></exception>
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
        ///     Returns <see cref="Data" /> as a string when <see cref="Type" /> is set to <see cref="MsgReader.Outlook.PropertyType.PT_UNICODE" />
        ///     or <see cref="MsgReader.Outlook.PropertyType.PT_STRING8" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_UNICODE"/> 
        /// or <see cref="MsgReader.Outlook.PropertyType.PT_STRING8" /></exception>
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
                        var encoding2 = Type == PropertyType.PT_STRING8 ? Encoding.Default : Encoding.Unicode;
                        using (var memoryStream = new MemoryStream(Data))
                        using (var streamReader = new StreamReader(memoryStream, encoding2))
                        {
                            var streamContent = streamReader.ReadToEnd();
                            return streamContent.TrimEnd('\0');
                        }
                        //throw new MRInvalidProperty("Type is not PT_UNICODE or PT_STRING8");
                }
            }
        }

        /// <summary>
        ///     Returns <see cref="Data" /> as a byte[] when <see cref="Type" /> is set to <see cref="MsgReader.Outlook.PropertyType.PT_BINARY" />
        ///     <see cref="MsgReader.Outlook.PropertyType.PT_OBJECT" />
        /// </summary>
        /// <exception cref="MRInvalidProperty">Raised when the <see cref="Type"/> is not set to <see cref="MsgReader.Outlook.PropertyType.PT_BINARY"/></exception>
        public byte[] ToBinary
        {
            get
            {
                switch (Type)
                {
                    case PropertyType.PT_BINARY:
                        return Data;

                    default:
                        throw new MRInvalidProperty("Type is not PT_BINARY");
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
        /// <param name="type">The <see cref="MsgReader.Outlook.PropertyType" /></param>
        /// <param name="data">The property data</param>
        public SProperty(ushort id, PropertyType type, byte[] data)
        {
            Id = id;
            Type = type;
            Data = data;
        }
        #endregion
    }
}
