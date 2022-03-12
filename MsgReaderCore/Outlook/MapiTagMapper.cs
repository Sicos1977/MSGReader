//
// MapiTagMapper.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2022 Magic-Sessions. (www.magic-sessions.com)
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
using System.Globalization;
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class that is used as a placeholder for the found mapped named mapi tags
        /// </summary>
        private class MapiTagMapping
        {
            #region Properties
            /// <summary>
            /// Contains the named property identifier
            /// </summary>
            public string PropertyIdentifier { get; }

            /// <summary>
            /// Contains the identifier that is found in the entry or string stream
            /// </summary>
            public string EntryOrStringIdentifier { get; }

            /// <summary>
            /// Returns <c>true</c> when an string identifier has been found in the string stream
            /// </summary>
            public bool HasStringIdentifier { get; }
            #endregion

            #region Constructor
            internal MapiTagMapping(string propertyIdentifier, string entryOrStringIdentifier, bool hasStringIdentifier = false)
            {
                PropertyIdentifier = propertyIdentifier;
                EntryOrStringIdentifier = entryOrStringIdentifier;
                HasStringIdentifier = hasStringIdentifier;
            }
            #endregion
        }

        /// <summary>
        /// Class used to map known MAPI tags to the internal used values
        /// </summary>
        private class MapiTagMapper : Storage
        {
            #region Constructor
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.MapiTagMapper" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal MapiTagMapper(Storage message) : base(message._rootStorage)
            {
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
            }
            #endregion

            #region GetMapping
            /// <summary>
            /// Returns a dictionary with all the property mappings
            /// </summary>
            /// <param name="propertyIdents">List with all the named property idents, e.g 8005, 8006, 801C, etc...</param>
            /// <returns></returns>
            internal List<MapiTagMapping> GetMapping(IEnumerable<string> propertyIdents)
            {
                var result = new List<MapiTagMapping>();
                var entryStreamBytes = GetStreamBytes(MapiTags.EntryStream);
                var stringStreamBytes = GetStreamBytes(MapiTags.StringStream);

                if (entryStreamBytes.Length == 0)
                    return result;

                foreach (var propertyIdent in propertyIdents)
                {
                    try
                    {
                        // To read the correct mapped property we need to calculate the offset in the entry stream
                        // The offset is calculated bij subtracting 32768 (8000 hex) from the named property and
                        // multiply the outcome with 8
                        var identValue = ushort.Parse(propertyIdent, NumberStyles.HexNumber);
                        var entryOffset = (identValue - 32768)*8;
                        if (entryOffset > entryStreamBytes.Length) continue;

                        string entryIdentString;

                        // We need the first 2 bytes for the mapping, but because the nameStreamBytes is in little 
                        // endian we need to swap the first 2 bytes
     
                        if (entryStreamBytes[entryOffset + 1] == 0)
                        {
                            var entryIdent = new[] {entryStreamBytes[entryOffset]};
                            entryIdentString = BitConverter.ToString(entryIdent).Replace("-", string.Empty);
                        }
                        else
                        {
                            var entryIdent = new[] { entryStreamBytes[entryOffset + 1], entryStreamBytes[entryOffset] };
                            entryIdentString = BitConverter.ToString(entryIdent).Replace("-", string.Empty);    
                        }

                        var stringOffset = ushort.Parse(entryIdentString, NumberStyles.HexNumber);

                        if (stringOffset >= stringStreamBytes.Length)
                        {
                            //Debug.Print($"{propertyIdent} - {entryIdentString}");
                            result.Add(new MapiTagMapping(propertyIdent, entryIdentString));
                            continue;
                        }

                        // Read the first 4 bytes to determine the length of the string to read
                        var stringLength = 0;

                        var len = stringStreamBytes.Length - stringOffset;

                        if (len == 1)
                        {
                            var bytes = new byte[1];
                            Buffer.BlockCopy(stringStreamBytes, stringOffset, bytes, 0, len);
                            stringLength = bytes[0];
                        }

                        if (len == 2)
                            stringLength = BitConverter.ToInt16(stringStreamBytes, stringOffset);

                        if (len == 3)
                        {
                            var bytes = new byte[3];
                            Buffer.BlockCopy(stringStreamBytes, stringOffset, bytes, 0, len);
                            stringLength = Bytes2Int(bytes[2], bytes[1], bytes[0]);
                        }
                        else if (len >= 4)
                            stringLength = BitConverter.ToInt32(stringStreamBytes, stringOffset);

                        if (stringOffset + stringLength >= stringStreamBytes.Length)
                        {
                            result.Add(new MapiTagMapping(propertyIdent, entryIdentString));
                            continue;
                        }

                        var str = string.Empty;

                        // Skip 4 bytes and start reading the string
                        stringOffset += 4;
                        for (var i = stringOffset; i < stringOffset + stringLength; i += 2)
                        {
                            var chr = BitConverter.ToChar(stringStreamBytes, i);
                            str += chr;
                        }

                        // Remove any null character
                        str = str.Replace("\0", string.Empty);
                        result.Add(new MapiTagMapping(str, propertyIdent, true));
                    }
                    catch (Exception exception)
                    {
                        Logger.WriteToLog(ExceptionHelpers.GetInnerException(exception));
                        throw;
                    }
                }

                return result;
            }
            #endregion

            #region Bytes2Int
            private static int Bytes2Int(byte b1, byte b2, byte b3)
            {
                int r = 0;
                byte b0 = 0xff;

                if ((b1 & 0x80) != 0) r |= b0 << 24;
                r |= b1 << 16;
                r |= b2 << 8;
                r |= b3;
                return r;
            }
            #endregion
        }
    }
}