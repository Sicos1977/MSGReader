using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;
// ReSharper disable LocalizableElement
// ReSharper disable UseNullPropagation
// ReSharper disable ConvertPropertyToExpressionBody
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Local
// ReSharper disable UseNameofExpression
// ReSharper disable MergeConditionalExpression

/*
   Copyright 2013-2017 Kees van Spelde

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
    /// <summary>
    /// The base class for reading an Outlook MSG file
    /// </summary>
    public partial class Storage : IDisposable
    {
        #region Fields
        /// <summary>
        /// The statistics for all streams in the IStorage associated with this instance
        /// </summary>
        private readonly Dictionary<string, STATSTG> _streamStatistics = new Dictionary<string, STATSTG>();

        /// <summary>
        /// The statistics for all storgages in the IStorage associated with this instance
        /// </summary>
        private readonly Dictionary<string, STATSTG> _subStorageStatistics = new Dictionary<string, STATSTG>();

        /// <summary>
        /// Header size of the property stream in the IStorage associated with this instance
        /// </summary>
        private int _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

        /// <summary>
        /// A reference to the parent message that this message may belong to
        /// </summary>
        private Storage _parentMessage;

        /// <summary>
        /// The IStorage associated with this instance.
        /// </summary>
        private NativeMethods.IStorage _storage;

        /// <summary>
        /// Will contain all the named MAPI properties when the class that inherits the <see cref="Storage"/> class 
        /// is a <see cref="Storage.Message"/> class. Otherwhise the List will be null
        /// mapped to
        /// </summary>
        private List<MapiTagMapping> _namedProperties;
        #endregion

        #region Properties
        /// <summary>
        /// Gets the top level Outlook message from a sub message at any level.
        /// </summary>
        /// <value> The top level outlook message. </value>
        private Storage TopParent
        {
            get { return _parentMessage != null ? _parentMessage.TopParent : this; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is the top level Outlook message.
        /// </summary>
        /// <value> <c>true</c> if this instance is the top level outlook message; otherwise, <c>false</c> . </value>
        private bool IsTopParent
        {
            get { return _parentMessage == null; }
        }

        /// <summary>
        /// The way the storage is opened
        /// </summary>
        public FileAccess FileAccess { get; private set; }
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
        /// <param name="fileAccess">FileAcces mode, default is Read</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        private Storage(string storageFilePath, FileAccess fileAccess = FileAccess.Read)
        {
            // Ensure provided file is an IStorage
            if (NativeMethods.StgIsStorageFile(storageFilePath) != 0)
                // ReSharper disable once LocalizableElement
                throw new ArgumentException("The provided file is not a valid IStorage", "storageFilePath");

            var accesMode = NativeMethods.STGM.READWRITE;
            FileAccess = fileAccess;

            switch (fileAccess)
            {
                case FileAccess.Read:
                    accesMode = NativeMethods.STGM.READ;
                    break;

                case FileAccess.Write:
                case FileAccess.ReadWrite:
                    accesMode = NativeMethods.STGM.READWRITE;
                    break;
            }

            // Open and load IStorage from file
            NativeMethods.IStorage fileStorage;
            NativeMethods.StgOpenStorage(storageFilePath, null,
                accesMode | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0, out fileStorage);

            // ReSharper disable once DoNotCallOverridableMethodsInConstructor
            LoadStorage(fileStorage);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class from a <see cref="Stream" /> containing an IStorage.
        /// </summary>
        /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
        /// <param name="fileAccess">FileAcces mode, default is Read</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        private Storage(Stream storageStream, FileAccess fileAccess = FileAccess.Read)
        {
            NativeMethods.IStorage memoryStorage = null;
            NativeMethods.ILockBytes memoryStorageBytes = null;
            try
            {
                // Read stream into buffer
                var buffer = new byte[storageStream.Length];
                storageStream.Read(buffer, 0, buffer.Length);

                // Create a ILockBytes (unmanaged byte array) and write buffer into it
                NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                memoryStorageBytes.WriteAt(0, buffer, buffer.Length, null);

                // Ensure provided stream data is an IStorage
                if (NativeMethods.StgIsStorageILockBytes(memoryStorageBytes) != 0)
                    // ReSharper disable once LocalizableElement
                    throw new ArgumentException("The provided stream is not a valid IStorage", "storageStream");

                var accesMode = NativeMethods.STGM.READWRITE;
                FileAccess = fileAccess;

                switch (fileAccess)
                {
                    case FileAccess.Read:
                        accesMode = NativeMethods.STGM.READ;
                        break;

                    case FileAccess.Write:
                    case FileAccess.ReadWrite:
                        accesMode = NativeMethods.STGM.READWRITE;
                        break;
                }

                // Open and load IStorage on the ILockBytes
                NativeMethods.StgOpenStorageOnILockBytes(memoryStorageBytes, null,
                    accesMode | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0, out memoryStorage);

                // ReSharper disable once DoNotCallOverridableMethodsInConstructor
                LoadStorage(memoryStorage);
            }
            catch
            {
                if (memoryStorage != null)
                    Marshal.ReleaseComObject(memoryStorage);

                throw;
            }
            finally
            {
                if (memoryStorageBytes != null)
                    Marshal.ReleaseComObject(memoryStorageBytes);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class on the specified <see cref="NativeMethods.IStorage" />.
        /// </summary>
        /// <param name="storage"> The storage to create the <see cref="Storage" /> on. </param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        private Storage(NativeMethods.IStorage storage)
        {
            // ReSharper disable once DoNotCallOverridableMethodsInConstructor
            LoadStorage(storage);
        }

        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="Storage" /> is reclaimed by garbage collection.
        /// </summary>
        ~Storage()
        {
            Dispose(false);
        }
        #endregion

        #region LoadStorage
        /// <summary>
        /// Processes sub streams and storages on the specified storage.
        /// </summary>
        /// <param name="storage"> The storage to get sub streams and storages for. </param>
        protected virtual void LoadStorage(NativeMethods.IStorage storage)
        {
            if (storage == null)
                throw new ArgumentNullException("storage", "Storage can not be null"); 
            
            _storage = storage;

            // Ensures memory is released
            ReferenceManager.AddItem(storage);
            NativeMethods.IEnumSTATSTG storageElementEnum = null;

            try
            {
                // Enum all elements of the storage
                storage.EnumElements(0, IntPtr.Zero, 0, out storageElementEnum);

                // Iterate elements
                while (true)
                {
                    // Get 1 element out of the COM enumerator
                    uint elementStatCount;
                    var elementStats = new STATSTG[1];
                    storageElementEnum.Next(1, elementStats, out elementStatCount);

                    // Break loop if element not retrieved
                    if (elementStatCount != 1)
                        break;

                    var elementStat = elementStats[0];
                    switch (elementStat.type)
                    {
                        case 1:
                            // Element is a storage, add its statistics object to the storage dictionary
                            _subStorageStatistics.Add(elementStat.pwcsName, elementStat);
                            break;

                        case 2:
                            // Element is a stream, add its statistics object to the stream dictionary
                            _streamStatistics.Add(elementStat.pwcsName, elementStat);
                            break;
                    }
                }
            }
            finally
            {
                // Free memory
                if (storageElementEnum != null)
                    Marshal.ReleaseComObject(storageElementEnum);
            }
        }
        #endregion

        #region GetStreamBytes
        /// <summary>
        /// Gets the data in the specified stream as a byte array. 
        /// Returns null when the <param ref="streamName"/> does not exists.
        /// </summary>
        /// <param name="streamName"> Name of the stream to get data for. </param>
        /// <returns> A byte array containg the stream data. </returns>
        private byte[] GetStreamBytes(string streamName)
        {
            if (!_streamStatistics.ContainsKey(streamName))
                return null;

            // Get statistics for stream 
            var streamStatStg = _streamStatistics[streamName];

            byte[] iStreamContent;
            IStream stream = null;
            try
            {
                // Open stream from the storage
                stream = _storage.OpenStream(streamStatStg.pwcsName, IntPtr.Zero,
                    NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, 0);

                // Read the stream into a managed byte array
                iStreamContent = new byte[streamStatStg.cbSize];
                stream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);
            }
            finally
            {
                if (stream != null)
                    Marshal.ReleaseComObject(stream);
            }

            // Return the stream bytes
            return iStreamContent;
        }
        #endregion

        #region GetStreamAsString
        /// <summary>
        /// Gets the data in the specified stream as a string using the specifed encoding to decode the stream data.
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

            var streamReader = new StreamReader(new MemoryStream(bytes), streamEncoding);
            var streamContent = streamReader.ReadToEnd();
            streamReader.Close();

            // Remove null termination chars when they exist
            return streamContent.Replace("\0", string.Empty);
        }
        #endregion

        #region GetMapiProperty
        /// <summary>
        /// Gets the raw value of the MAPI property.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The raw value of the MAPI property. </returns>
        private object GetMapiProperty(string propIdentifier)
        {
            // Check if the propIdentifier is a named property and if so replace it with
            // the correct mapped property
            if (_namedProperties != null)
            {
                var mapiTagMapping = _namedProperties.Find(m => m.EntryOrStringIdentifier == propIdentifier);
                if (mapiTagMapping != null)
                    propIdentifier = mapiTagMapping.PropertyIdentifier;
            }

            // Try get prop value from stream or storage
            // If not found in stream or storage try get prop value from property stream
            var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                            GetMapiPropertyFromPropertyStream(propIdentifier);

            return propValue;
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
                    //return GetStreamAsString(containerName, Encoding.UTF8);
                    return GetStreamAsString(containerName, Encoding.Default);

                case PropertyType.PT_UNICODE:
                    return GetStreamAsString(containerName, Encoding.Unicode);

                case PropertyType.PT_BINARY:
                    return GetStreamBytes(containerName);

                case PropertyType.PT_MV_STRING8:
                case PropertyType.PT_MV_UNICODE:

                    // If the property is a unicode multiview item we need to read all the properties
                    // again and filter out all the multivalue names, they end with -00000000, -00000001, etc..
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

                case PropertyType.PT_OBJECT:
                    return
                        NativeMethods.CloneStorage(
                            _storage.OpenStorage(containerName, IntPtr.Zero,
                                NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE,
                                IntPtr.Zero, 0), true);

                default:
                    throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
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

                // Get property identifer located in 3nd and 4th bytes as a hexdecimal string
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

                    //default:
                    //throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
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
            return list == null ? null : list.AsReadOnly();
        }

        /// <summary>
        /// Gets the value of the MAPI property as a integer.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a integer. </returns>
        private int? GetMapiPropertyInt32(string propIdentifier)
        {
            var value = GetMapiProperty(propIdentifier);

            if (value != null)
                return (int)value;

            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a double.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a double. </returns>
        private double? GetMapiPropertyDouble(string propIdentifier)
        {
            var value = GetMapiProperty(propIdentifier);

            if (value != null)
                return (double)value;

            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a datetime.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a datetime or null when not set </returns>
        private DateTime? GetMapiPropertyDateTime(string propIdentifier)
        {
            var value = GetMapiProperty(propIdentifier);

            if (value != null)
                return (DateTime)value;

            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a bool.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
        /// <returns> The value of the MAPI property as a boolean or null when not set. </returns>
        private bool? GetMapiPropertyBool(string propIdentifier)
        {
            var value = GetMapiProperty(propIdentifier);

            if (value != null)
                return (bool)value;

            return null;
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
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposes this object
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
                Disposing();

            if (_storage == null) return;
            ReferenceManager.RemoveItem(_storage);
            Marshal.ReleaseComObject(_storage);
            _storage = null;
        }

        /// <summary>
        /// Gives sub classes the chance to free resources during object disposal.
        /// </summary>
        protected virtual void Disposing()
        {
        }
        #endregion
    }
}