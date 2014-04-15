using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {
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
            internal MapiTagMapper(Storage message) : base(message._storage)
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
            internal Dictionary<string, string> GetMapping(IEnumerable<string> propertyIdents)
            {
                var result = new Dictionary<string, string>();
                var entryStreamBytes = GetStreamBytes(MapiTags.EntryStream);
                var stringStreamBytes = GetStreamBytes(MapiTags.StringStream);

                foreach (var propertyIdent in propertyIdents)
                {
                    // To read the correct mapped property we need to calculate the ofset in the entry stream
                    // The offset is calculated bij substracting 32768 (8000 hex) from the named property and
                    // multiply the outcome with 8
                    var identValue = ushort.Parse(propertyIdent, NumberStyles.HexNumber);
                    var entryOffset = (identValue - 32768)*8;

                    // We need the first 2 bytes for the mapping, but because the nameStreamBytes is in little 
                    // endian we need to swap the first 2 bytes
                    var entryIdent = new[] {entryStreamBytes[entryOffset + 1], entryStreamBytes[entryOffset]};
                    var entryIdentString = BitConverter.ToString(entryIdent).Replace("-", string.Empty);

                    // When the type = 05 it means we have to look for a mapping in the string stream
                    // 03-E8-00-00-05-00-FE-00
                    var type = BitConverter.ToString(entryStreamBytes, entryOffset + 4, 1);
                    if (type == "05")
                    {
                        var stringOffset = ushort.Parse(entryIdentString, NumberStyles.HexNumber);

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
                        result.Add(str, propertyIdent);
                    }
                    else
                    {
                        // Convert it to a short
                        var newIdentValue = ushort.Parse(entryIdentString, NumberStyles.HexNumber);

                        // Check if the value is in the named property range (8000 to FFFE (Hex))
                        if (newIdentValue >= 32768 && newIdentValue <= 65534)
                        //if (!result.ContainsKey(entryIdentString))
                            result.Add(entryIdentString, propertyIdent);
                    }
                }

                return result;
            }
            #endregion
        }
    }
}