//
// Storage.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2021 Magic-Sessions. (www.magic-sessions.com)
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
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using MsgReader.Helpers;
using OpenMcdf;
// ReSharper disable MemberCanBePrivate.Global

namespace MsgReader.Outlook
{
    /// <summary>
    /// The base class for reading an Outlook MSG file
    /// </summary>
    public partial class Storage : IDisposable
    {
        #region Fields
        /// <summary>
        /// The statistics for all streams in the Storage associated with this instance
        /// </summary>
        private readonly Dictionary<string, CFStream> _streamStatistics = new Dictionary<string, CFStream>();

        /// <summary>
        /// The statistics for all storage in the Storage associated with this instance
        /// </summary>
        private readonly Dictionary<string, CFStorage> _subStorageStatistics = new Dictionary<string, CFStorage>();

        /// <summary>
        /// Header size of the property stream in the IStorage associated with this instance
        /// </summary>
        private int _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

        /// <summary>
        /// A reference to the parent message that this message may belong to
        /// </summary>
        private Storage _parentMessage;

        /// <summary>
        /// The opened compound file
        /// </summary>
        private CompoundFile _compoundFile;

        /// <summary>
        /// The root storage associated with this instance.
        /// </summary>
        private CFStorage _rootStorage;

        /// <summary>
        /// Will contain all the named MAPI properties when the class that inherits the <see cref="Storage"/> class 
        /// is a <see cref="Storage.Message"/> class. Otherwise the List will be null
        /// mapped to
        /// </summary>
        private List<MapiTagMapping> _namedProperties;

        /// <summary>
        /// Returns <c>true</c> when one or more attachment are deleted
        /// </summary>
        private bool _attachmentDeleted;

        /// <summary>
        /// Contains the <see cref="Encoding"/> that is used for the <see cref="Message.BodyText"/> or <see cref="Message.BodyHtml"/>. 
        /// It will contain null when the codepage could not be read from the <see cref="Storage.Message"/>
        /// </summary>
        private Encoding _internetCodepage;

        /// <summary>
        /// Contains the <see cref="Encoding"/> that is used for the <see cref="Message.BodyRtf"/>.
        /// It will contain null when the codepage could not be read from the <see cref="Storage.Message"/>
        /// </summary>
        private Encoding _messageCodepage;
        #endregion

        #region Properties
        /// <summary>
        /// Gets the top level Outlook message from a sub message at any level.
        /// </summary>
        /// <value> The top level Outlook message. </value>
        private Storage TopParent => _parentMessage != null ? _parentMessage.TopParent : this;

        /// <summary>
        /// Gets a value indicating whether this instance is the top level Outlook message.
        /// </summary>
        /// <value> <c>true</c> if this instance is the top level Outlook message; otherwise, <c>false</c> . </value>
        private bool IsTopParent => _parentMessage == null;
        
        /// <summary>
        /// The way the storage is opened
        /// </summary>
        public FileAccess FileAccess { get; }

        /// <summary>
        /// Returns the <see cref="Encoding"/> that is used for the <see cref="Message.BodyText"/>
        /// or <see cref="Message.BodyHtml"/>
        /// <remarks>
        /// See the <see cref="MessageCodePage"/> property when dealing with the <see cref="Message.BodyRtf"/>
        /// </remarks>
        /// </summary>
        public Encoding InternetCodePage
        {
            get
            {
                if (_internetCodepage != null)
                    return _internetCodepage;

                var codePage = GetMapiPropertyInt32(MapiTags.PR_INTERNET_CPID);
                _internetCodepage = codePage == null ? Encoding.Default : Encoding.GetEncoding((int)codePage);
                return _internetCodepage;
            }
        }

        /// <summary>
        /// Returns the <see cref="Encoding"/> that is used for the <see cref="Message.BodyRtf"/>.
        /// It will return the systems default encoding when the codepage could not be read from 
        /// the <see cref="Storage.Message"/>
        /// <remarks>
        /// See the <see cref="InternetCodePage"/> property when dealing with the <see cref="Message.BodyRtf"/>
        /// </remarks>
        /// </summary>
        public Encoding MessageCodePage
        {
            get
            {
                if (_messageCodepage != null)
                    return _messageCodepage;

                var codePage = GetMapiPropertyInt32(MapiTags.PR_MESSAGE_CODEPAGE);

                try
                {
                    _messageCodepage = codePage != null ? Encoding.GetEncoding((int)codePage) : InternetCodePage;
                }
                catch (NotSupportedException)
                {
                    _messageCodepage = InternetCodePage;
                }

                return _messageCodepage;
            }
        }
        #endregion

        #region Constructors & Destructor
        // ReSharper disable once UnusedMember.Local
        private Storage()
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class from a file.
        /// </summary>
        /// <param name="storageFilePath"> The file to load. </param>
        /// <param name="fileAccess">FileAccess mode, default is Read</param>
        private Storage(string storageFilePath, FileAccess fileAccess = FileAccess.Read)
        {
            FileAccess = fileAccess;

            switch(FileAccess)
            {
                case FileAccess.Read:
                    _compoundFile = new CompoundFile(storageFilePath);
                    break;
                case FileAccess.Write:
                case FileAccess.ReadWrite:
                    _compoundFile = new CompoundFile(storageFilePath, CFSUpdateMode.Update, CFSConfiguration.Default);
                    break;
            }

            // ReSharper disable once VirtualMemberCallInConstructor
            LoadStorage(_compoundFile.RootStorage);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class from a <see cref="Stream" /> containing an IStorage.
        /// </summary>
        /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
        /// <param name="fileAccess">FileAccess mode, default is Read</param>
        private Storage(Stream storageStream, FileAccess fileAccess = FileAccess.Read)
        {
            FileAccess = fileAccess;

            switch (FileAccess)
            {
                case FileAccess.Read:
                    _compoundFile = new CompoundFile(storageStream);
                    break;
                case FileAccess.Write:
                case FileAccess.ReadWrite:
                    _compoundFile = new CompoundFile(storageStream, CFSUpdateMode.Update, CFSConfiguration.Default);
                    break;
            }

            // ReSharper disable once VirtualMemberCallInConstructor
            LoadStorage(_compoundFile.RootStorage);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class on the specified <see cref="CFStorage" />.
        /// </summary>
        /// <param name="storage"> The storage to create the <see cref="Storage" /> on. </param>
        private Storage(CFStorage storage)
        {
            // ReSharper disable once VirtualMemberCallInConstructor
            LoadStorage(storage);
        }

        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="Storage" /> is reclaimed by garbage collection.
        /// </summary>
        ~Storage()
        {
            Dispose();
        }
        #endregion

        #region LoadStorage
        /// <summary>
        /// Processes sub streams and storage on the specified storage.
        /// </summary>
        /// <param name="storage"> The storage to get sub streams and storage for. </param>
        protected virtual void LoadStorage(CFStorage storage)
        {
            if (storage == null) return;

            _rootStorage = storage;

            storage.VisitEntries(cfItem =>
            {
                if (cfItem.IsStorage)
                    _subStorageStatistics.Add(cfItem.Name, cfItem as CFStorage);
                else
                    _streamStatistics.Add(cfItem.Name, cfItem as CFStream);

            }, false);
        }
        #endregion

        #region GetStreamBytes
        /// <summary>
        /// Gets the data in the specified stream as a byte array. 
        /// Returns null when the <param ref="streamName"/> does not exists.
        /// </summary>
        /// <param name="streamName"> Name of the stream to get data for. </param>
        /// <returns> A byte array containing the stream data. </returns>
        private byte[] GetStreamBytes(string streamName)
        {
            if (!_streamStatistics.ContainsKey(streamName))
                return null;

            // Get statistics for stream 
            var stream = _streamStatistics[streamName];
            return stream.GetData();
        }

        /// <summary>
        /// Returns a <see cref="Stream"/> for the given <paramref name="streamName"/>
        /// <c>null</c> is returned when the stream does not exists
        /// </summary>
        /// <param name="streamName"></param>
        /// <returns></returns>
        private Stream GetStream(string streamName)
        {
            if (!_streamStatistics.ContainsKey(streamName))
                return null;

            // Get statistics for stream 
            var stream = _streamStatistics[streamName];
            return new MemoryStream(stream.GetData());
        }
        #endregion

        #region GetStreamAsString
        /// <summary>
        /// Gets the data in the specified stream as a string using the specified encoding to decode the stream data.
        /// Returns null when the <param ref="streamName"/> does not exists.
        /// </summary>
        /// <param name="streamName"> Name of the stream to get string data for. </param>
        /// <param name="streamEncoding"> The encoding to decode the stream data with. </param>
        /// <returns> The data in the specified stream as a string. </returns>
        private string GetStreamAsString(string streamName, Encoding streamEncoding)
        {
            var bytes = GetStreamBytes(streamName);
            if (bytes == null)
                return null;

            using (var streamReader = new StreamReader(new MemoryStream(bytes), streamEncoding))
            {
                var streamContent = streamReader.ReadToEnd();
                // Remove null termination chars when they exist
                return streamContent.Replace("\0", string.Empty);
            }
        }
        #endregion

        #region GetMapiProperty
        /// <summary>
        /// Gets the raw value of the MAPI property.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The raw value of the MAPI property. </returns>
        public object GetMapiProperty(string propIdentifier)
        {
            Logger.WriteToLog($"Getting property with id '{propIdentifier}'");

            // Check if the propIdentifier is a named property and if so replace it with
            // the correct mapped property
            var mapiTagMapping = _namedProperties?.Find(m => m.PropertyIdentifier == propIdentifier);
            if (mapiTagMapping != null)
                propIdentifier = mapiTagMapping.EntryOrStringIdentifier;

            // Try get prop value from stream or storage
            // If not found in stream or storage try get prop value from property stream
            var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                            GetMapiPropertyFromPropertyStream(propIdentifier);

            return propValue;
        }
        #endregion

        #region GetNamedProperties
        /// <summary>
        /// Returns a list with all the found named properties
        /// </summary>
        /// <returns></returns>
        // ReSharper disable once UnusedMember.Global
        public List<string> GetNamedProperties()
        {
            return (from namedProperty in _namedProperties
                where namedProperty.HasStringIdentifier
                select namedProperty.PropertyIdentifier).ToList();
        }
        #endregion

        #region GetNamedMapiProperty
        /// <summary>
        /// Gets the raw value of the MAPI property.
        /// </summary>
        /// <param name="propertyIdentifier">The name of the property</param>
        /// <returns>The raw value of the MAPI property or <c>null</c> when not found</returns>
        // ReSharper disable once UnusedMember.Global
        public object GetNamedMapiProperty(string propertyIdentifier)
        {
            Logger.WriteToLog($"Getting named property with name '{propertyIdentifier}'");

            // Check if the propIdentifier is a named property and if so replace it with
            // the correct mapped property
            var mapiTagMapping = _namedProperties?.Find(m => m.PropertyIdentifier == propertyIdentifier && m.HasStringIdentifier);
            if (mapiTagMapping != null)
            {
                var propIdentifier = mapiTagMapping.EntryOrStringIdentifier;

                // Try get prop value from stream or storage
                // If not found in stream or storage try get prop value from property stream
                var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                                GetMapiPropertyFromPropertyStream(propIdentifier);

                return propValue;
            }

            return null;
        }
        #endregion

        #region GetMapiPropertyFromStreamOrStorage
        /// <summary>
        /// Gets the MAPI property value from a stream or storage in this storage.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property or null if not found. </returns>
        private object GetMapiPropertyFromStreamOrStorage(string propIdentifier)
        {
            // Get list of stream and storage identifiers which map to properties
            var propKeys = new List<string>();
            propKeys.AddRange(_streamStatistics.Keys);
            propKeys.AddRange(_subStorageStatistics.Keys);

            // Determine if the property identifier is in a stream or sub storage
            string propTag = null;
            var propType = PropertyType.PT_UNSPECIFIED;

            foreach (var propKey in propKeys)
            {
                if (!propKey.StartsWith(MapiTags.SubStgVersion1 + "_" + propIdentifier)) continue;
                propTag = propKey.Substring(12, 8);
                propType = (PropertyType) ushort.Parse(propKey.Substring(16, 4), NumberStyles.HexNumber);
                break;
            }

            // When null then we didn't find the property
            if (propTag == null)
                return null;

            // Depending on prop type use method to get property value
            var containerName = MapiTags.SubStgVersion1 + "_" + propTag;
            switch (propType)
            {
                case PropertyType.PT_UNSPECIFIED:
                    return null;

                case PropertyType.PT_STRING8:
                    return GetStreamAsString(containerName, Encoding.Default);

                case PropertyType.PT_UNICODE:
                    return GetStreamAsString(containerName, Encoding.Unicode);

                case PropertyType.PT_BINARY:
                    return GetStreamBytes(containerName);

                case PropertyType.PT_MV_STRING8:
                case PropertyType.PT_MV_UNICODE:

                    // If the property is a unicode multi view item we need to read all the properties
                    // again and filter out all the multi value names, they end with -00000000, -00000001, etc..
                    var multiValueContainerNames = propKeys.Where(propKey => propKey.StartsWith(containerName + "-")).ToList();

                    var values = new List<string>();
                    foreach (var multiValueContainerName in multiValueContainerNames)
                    {
                        var value = GetStreamAsString(multiValueContainerName,
                            propType == PropertyType.PT_MV_STRING8 ? Encoding.Default : Encoding.Unicode);

                        // Multi values always end with a null char so we need to strip that one off
                        if (value.EndsWith("/0"))
                            value = value.Substring(0, value.Length - 1);

                        values.Add(value);
                    }

                    return values;

                case PropertyType.PT_MV_LONG:
                    using (var binaryReader = new BinaryReader(GetStream(containerName)))
                    {
                        var result = new List<long>();
                        while (binaryReader.PeekChar() != -1)
                        {
                            var value = binaryReader.ReadUInt16();
                            if (value > 0)
                                result.Add(value);
                        }

                        return result;
                    }

                case PropertyType.PT_OBJECT:
                    return _subStorageStatistics[containerName];

                default:
                    throw new ApplicationException($"MAPI property has an unsupported type '{propType}' and can not be retrieved.");
            }
        }

        /// <summary>
        /// Gets the MAPI property value from the property stream in this storage.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property or null if not found. </returns>
        private object GetMapiPropertyFromPropertyStream(string propIdentifier)
        {
            // If no property stream return null
            if (!_streamStatistics.ContainsKey(MapiTags.PropertiesStream))
                return null;

            // Get the raw bytes for the property stream
            var propBytes = GetStreamBytes(MapiTags.PropertiesStream);

            // Iterate over property stream in 16 byte chunks starting from end of header
            for (var i = _propHeaderSize; i < propBytes.Length; i = i + 16)
            {
                // Get property type located in the 1st and 2nd bytes as a unsigned short value
                var propType = (PropertyType) BitConverter.ToUInt16(propBytes, i);

                // Get property identifier located in 3nd and 4th bytes as a hexadecimal string
                var propIdent = new[] { propBytes[i + 3], propBytes[i + 2] };
                var propIdentString = BitConverter.ToString(propIdent).Replace("-", string.Empty);

                // If this is not the property being gotten continue to next property
                if (propIdentString != propIdentifier) continue;

                // Depending on prop type use method to get property value
                switch (propType)
                {
                    case PropertyType.PT_SHORT:
                        return BitConverter.ToInt16(propBytes, i + 8);

                    case PropertyType.PT_LONG:
                        return BitConverter.ToInt32(propBytes, i + 8);

                    case PropertyType.PT_DOUBLE:
                        return BitConverter.ToDouble(propBytes, i + 8);

                    case PropertyType.PT_SYSTIME:
                        var fileTime = BitConverter.ToInt64(propBytes, i + 8);
                        return DateTime.FromFileTime(fileTime);

                    case PropertyType.PT_APPTIME:
                        var appTime = BitConverter.ToInt64(propBytes, i + 8);
                        return DateTime.FromOADate(appTime);

                    case PropertyType.PT_BOOLEAN:
                        return BitConverter.ToBoolean(propBytes, i + 8);
                }
            }

            // Property not found return null
            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as an <see cref="UnsendableRecipients"/>
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a string. </returns>
        private UnsendableRecipients GetUnsendableRecipients(string propIdentifier)
        {
            var data = GetMapiPropertyBytes(propIdentifier);
            return data != null ? new UnsendableRecipients(data) : null;
        } 

        /// <summary>
        /// Gets the value of the MAPI property as a string.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a string. </returns>
        private string GetMapiPropertyString(string propIdentifier)
        {
            return GetMapiProperty(propIdentifier) as string;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a list of string.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a list of string. </returns>
        private ReadOnlyCollection<string> GetMapiPropertyStringList(string propIdentifier)
        {
            var list = GetMapiProperty(propIdentifier) as List<string>;
            return list?.AsReadOnly();
        }

        /// <summary>
        /// Gets the value of the MAPI property as a integer.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a integer. </returns>
        private int? GetMapiPropertyInt32(string propIdentifier)
        {
            return (int?)GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a double.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a double. </returns>
        private double? GetMapiPropertyDouble(string propIdentifier)
        {
            return (double?) GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a datetime.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a datetime or null when not set </returns>
        private DateTime? GetMapiPropertyDateTime(string propIdentifier)
        {
            return (DateTime?)GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a bool.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a boolean or null when not set. </returns>
        private bool? GetMapiPropertyBool(string propIdentifier)
        {
            return (bool?)GetMapiProperty(propIdentifier); 
        }

        /// <summary>
        /// Gets the value of the MAPI property as a byte array.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a byte array. </returns>
        private byte[] GetMapiPropertyBytes(string propIdentifier)
        {
            return (byte[])GetMapiProperty(propIdentifier);
        }
        #endregion

        #region IDisposable Members
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_compoundFile == null) return;
            try
            {
                _compoundFile.Close();
            }
            catch 
            {
                // Ignore
            }
            _compoundFile = null;
        }
        #endregion
    }
}