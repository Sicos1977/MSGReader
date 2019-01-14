//
// MapiTagMapper.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
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

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class that is used as a placeholder for the found mapped named mapi tags
        /// </summary>
        internal class MapiTagMapping
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
            #endregion

            #region Constructor
            internal MapiTagMapping(string propertyIdentifier, string entryOrStringIdentifier)
            {
                PropertyIdentifier = propertyIdentifier;
                EntryOrStringIdentifier = entryOrStringIdentifier;
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
                GC.SuppressFinalize(message);
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
                    // To read the correct mapped property we need to calculate the offset in the entry stream
                    // The offset is calculated bij substracting 32768 (8000 hex) from the named property and
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

                    // When the type = 05 it means we have to look for a mapping in the string stream
                    // 03-E8-00-00-05-00-FE-00
                    var type = BitConverter.ToString(entryStreamBytes, entryOffset + 4, 1);
                    if (type == "05")
                    {
                        var stringOffset = ushort.Parse(entryIdentString, NumberStyles.HexNumber);

                        if (stringOffset >= stringStreamBytes.Length) continue;
                        // Read the first 4 bytes to determine the length of the string to read
                        var stringLength = BitConverter.ToInt32(stringStreamBytes, stringOffset);
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
                        result.Add(new MapiTagMapping(propertyIdent, str));
                    }
                    else
                        result.Add(new MapiTagMapping(propertyIdent, entryIdentString));
                }

                return result;
            }
            #endregion
        }
    }
}