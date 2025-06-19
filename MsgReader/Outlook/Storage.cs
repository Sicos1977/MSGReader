﻿//
// Storage.cs
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
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using MsgReader.Helpers;
using OpenMcdf;

// ReSharper disable MemberCanBePrivate.Global

namespace MsgReader.Outlook;

/// <summary>
///     The base class for reading an Outlook MSG file
/// </summary>
public partial class Storage : IDisposable
{
    #region Fields
    /// <summary>
    ///     The statistics for all streams in the Storage associated with this instance
    /// </summary>
    private readonly Dictionary<string, CfbStream> _streamStatistics;

    /// <summary>
    ///     The statistics for all storage in the Storage associated with this instance
    /// </summary>
    private readonly Dictionary<string, OpenMcdf.Storage> _subStorageStatistics;

    /// <summary>
    ///     Header size of the property stream in the IStorage associated with this instance
    /// </summary>
    private int _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

    /// <summary>
    ///     When set to <c>true</c> then the given <see cref="Stream" /> is not closed after use
    /// </summary>
    private readonly bool _leaveStreamOpen;

    /// <summary>
    ///     The root storage associated with this instance
    /// </summary>
    private RootStorage _rootStorage;

    ///// <summary>
    /////     The storage associated with this instance
    ///// </summary>
    private OpenMcdf.Storage _storage;

    /// <summary>
    ///     Will contain all the named MAPI properties when the class that inherits the <see cref="Storage" /> class
    ///     is a <see cref="Storage.Message" /> class. Otherwise, the List will be null
    ///     mapped to
    /// </summary>
    private List<MapiTagMapping> _namedProperties;

    /// <summary>
    ///     Contains the <see cref="Encoding" /> that is used for the <see cref="Message.BodyText" /> or
    ///     <see cref="Message.BodyHtml" />.
    ///     It will contain null when the codepage could not be read from the <see cref="Storage.Message" />
    /// </summary>
    private Encoding _internetCodepage;

    /// <summary>
    ///     <see cref="LocalId" />
    /// </summary>
    private int? _localId;

    /// <summary>
    ///     Contains the <see cref="Encoding" /> that is used for the <see cref="Message.BodyRtf" />.
    ///     It will contain null when the codepage could not be read from the <see cref="Storage.Message" />
    /// </summary>register
    private Encoding _messageCodepage;
    #endregion

    #region Properties
    /// <summary>
    ///     The way the storage is opened
    /// </summary>
    public FileAccess FileAccess { get; }

    /// <summary>
    ///     Returns the <see cref="Encoding" /> that is used for the <see cref="Message.BodyText" />
    ///     or <see cref="Message.BodyHtml" />
    ///     <remarks>
    ///         See the <see cref="MessageCodePage" /> property when dealing with the <see cref="Message.BodyRtf" />
    ///     </remarks>
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
    ///     Returns the Windows LCID of the end user who created this message
    /// </summary>
    public int? LocalId => _localId ??= GetMapiPropertyInt32(MapiTags.PR_MESSAGE_LOCALE_ID);

    /// <summary>
    ///     Returns the <see cref="Encoding" /> that is used for the <see cref="Message.BodyRtf" />.
    ///     It will return the systems default encoding when the codepage could not be read from
    ///     the <see cref="Storage.Message" />
    ///     <remarks>
    ///         See the <see cref="InternetCodePage" /> property when dealing with the <see cref="Message.BodyRtf" />
    ///     </remarks>
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
        _subStorageStatistics = new Dictionary<string, OpenMcdf.Storage>();
        _streamStatistics = new Dictionary<string, CfbStream>();
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Storage" /> class from a file.
    /// </summary>
    /// <param name="storageFilePath"> The file to load. </param>
    /// <param name="fileAccess">FileAccess mode, default is Read</param>
    private Storage(string storageFilePath, FileAccess fileAccess = FileAccess.Read)
    {
        _streamStatistics = new Dictionary<string, CfbStream>();
        _subStorageStatistics = new Dictionary<string, OpenMcdf.Storage>();
        FileAccess = fileAccess;

        switch(FileAccess)
        {
            case FileAccess.Read:
                _rootStorage = RootStorage.OpenRead(storageFilePath, StorageModeFlags.LeaveOpen);
                break;

            case FileAccess.Write:
            case FileAccess.ReadWrite:
                _rootStorage = RootStorage.Open(storageFilePath, FileMode.Open, fileAccess, StorageModeFlags.LeaveOpen | StorageModeFlags.Transacted);
                break;
        }

        // ReSharper disable once VirtualMemberCallInConstructor
        // ReSharper disable once PossibleNullReferenceException
        LoadStorage(_rootStorage);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Storage" /> class from a <see cref="Stream" /> containing an IStorage.
    /// </summary>
    /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
    /// <param name="fileAccess">FileAccess mode, default is Read</param>
    /// <param name="leaveStreamOpen">When set to <c>true</c> then the given <paramref name="storageStream"/> is not closed after use</param>
    private Storage(Stream storageStream, FileAccess fileAccess = FileAccess.Read, bool leaveStreamOpen = false)
    {
        _streamStatistics = new Dictionary<string, CfbStream>();
        _subStorageStatistics = new Dictionary<string, OpenMcdf.Storage>();
        FileAccess = fileAccess;
        _leaveStreamOpen = leaveStreamOpen;

        switch(FileAccess)
        {
            case FileAccess.Read:
                _rootStorage = RootStorage.Open(storageStream, leaveStreamOpen ? StorageModeFlags.LeaveOpen : StorageModeFlags.None);
                break;

            case FileAccess.Write:
            case FileAccess.ReadWrite:
                _rootStorage = RootStorage.Open(storageStream, leaveStreamOpen ? StorageModeFlags.LeaveOpen | StorageModeFlags.Transacted : StorageModeFlags.Transacted);
                break;
        }

        // ReSharper disable once VirtualMemberCallInConstructor
        // ReSharper disable once PossibleNullReferenceException
        LoadStorage(_rootStorage);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Storage" /> class on the specified <see cref="Storage" />.
    /// </summary>
    /// <param name="storage"> The storage to create the <see cref="Storage" /> on. </param>
    private Storage(OpenMcdf.Storage storage)
    {
        _subStorageStatistics = new Dictionary<string, OpenMcdf.Storage>();
        _streamStatistics = new Dictionary<string, CfbStream>();
        // ReSharper disable once VirtualMemberCallInConstructor
        LoadStorage(storage);
    }

    /// <summary>
    ///     Releases unmanaged resources and performs other cleanup operations before the
    ///     <see cref="Storage" /> is reclaimed by garbage collection.
    /// </summary>
    ~Storage()
    {
        Dispose();
    }
    #endregion

    #region LoadStorage
    /// <summary>
    ///     Processes sub streams and storage on the specified storage.
    /// </summary>
    /// <param name="storage"> The storage to get sub streams and storage for. </param>
    protected virtual void LoadStorage(OpenMcdf.Storage storage)
    {
        if (storage == null) return;

#if (NETSTANDARD2_0)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
        _storage = storage;

        foreach (var item in storage.EnumerateEntries())
        {
            if (item.Type == EntryType.Storage)
                _subStorageStatistics.Add(item.Name, storage.OpenStorage(item.Name));
            else
                _streamStatistics.Add(item.Name, storage.OpenStream(item.Name));
        }
    }
    #endregion

    #region GetStreamBytes
    /// <summary>
    ///     Gets the data in the specified stream as a byte array.
    ///     Returns null when the
    ///     <param ref="streamName" />
    ///     does not exist.
    /// </summary>
    /// <param name="streamName"> Name of the stream to get data for. </param>
    /// <returns> A byte array containing the stream data. </returns>
    private byte[] GetStreamBytes(string streamName)
    {
        if (!_streamStatistics.ContainsKey(streamName))
            return null;

        Logger.WriteToLog($"Getting stream with name '{streamName}' as bytes");

        // Get statistics for stream 
        var stream = _streamStatistics[streamName];
        using var memoryStream = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(memoryStream);
        return memoryStream.ToArray();
    }

    /// <summary>
    ///     Returns a <see cref="Stream" /> for the given <paramref name="streamName" />
    ///     <c>null</c> is returned when the stream does not exist
    /// </summary>
    /// <param name="streamName"></param>
    /// <returns></returns>
    private Stream GetStream(string streamName)
    {
        if (!_streamStatistics.ContainsKey(streamName))
            return null;

        Logger.WriteToLog($"Getting stream with name '{streamName}'");

        // Get statistics for stream 
        var data = _streamStatistics[streamName];
        var stream = StreamHelpers.Manager.GetStream("Storage.cs", data.Length);
        data.Position = 0;
        data.CopyTo(stream);
        return stream;
    }
    #endregion

    #region GetStreamAsString
    /// <summary>
    ///     Gets the data in the specified stream as a string using the specified encoding to decode the stream data.
    ///     Returns null when the
    ///     <param ref="streamName" />
    ///     does not exist.
    /// </summary>
    /// <param name="streamName"> Name of the stream to get string data for. </param>
    /// <param name="streamEncoding"> The encoding to decode the stream data with. </param>
    /// <returns> The data in the specified stream as a string. </returns>
    private string GetStreamAsString(string streamName, Encoding streamEncoding)
    {
        Logger.WriteToLog($"Getting stream with name '{streamName}' as string with encoding '{streamEncoding}'");

        var bytes = GetStreamBytes(streamName);
        if (bytes == null)
            return null;

        using var streamReader = new StreamReader(StreamHelpers.Manager.GetStream("Message.cs", bytes, 0, bytes.Length), streamEncoding);
        var streamContent = streamReader.ReadToEnd();
        // Remove null termination chars when they exist
        return streamContent.Replace("\0", string.Empty);
    }
    #endregion

    #region GetMapiProperty
    /// <summary>
    ///     Gets the raw value of the MAPI property.
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
        else
        {
            mapiTagMapping = _namedProperties?.Find(m => m.EntryOrStringIdentifier == propIdentifier);

            if (mapiTagMapping != null)
                propIdentifier = mapiTagMapping.PropertyIdentifier;
        }

        // Try to get prop value from stream or storage
        // If not found in stream or storage try to get prop value from property stream
        var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                        GetMapiPropertyFromPropertyStream(propIdentifier);

        return propValue;
    }
    #endregion

    #region GetNamedProperties
    /// <summary>
    ///     Returns a list with all the found named properties
    /// </summary>
    /// <returns></returns>
    // ReSharper disable once UnusedMember.Global
    public List<string> GetNamedProperties()
    {
        Logger.WriteToLog("Getting named properties");

        return (from namedProperty in _namedProperties
            where namedProperty.HasStringIdentifier
            select namedProperty.PropertyIdentifier).ToList();
    }
    #endregion

    #region GetNamedMapiProperty
    /// <summary>
    ///     Gets the raw value of the MAPI property.
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

            // Try to get prop value from stream or storage
            // If not found in stream or storage try ro get prop value from property stream
            var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                            GetMapiPropertyFromPropertyStream(propIdentifier);

            return propValue;
        }

        return null;
    }
    #endregion

    #region GetMapiPropertyFromStreamOrStorage
    /// <summary>
    ///     Gets the MAPI property value from a stream or storage in this storage.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
    /// <returns> The value of the MAPI property or null if not found. </returns>
    private object GetMapiPropertyFromStreamOrStorage(string propIdentifier)
    {
        Logger.WriteToLog($"Getting property with id '{propIdentifier}' from stream or storage");

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
            propType = (PropertyType)ushort.Parse(propKey.Substring(16, 4), NumberStyles.HexNumber);
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
                return GetStreamAsString(containerName, MessageCodePage);

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

                    // Multi values always end with a null char, so we need to strip that one off
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
    ///     Gets the MAPI property value from the property stream in this storage.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
    /// <returns> The value of the MAPI property or null if not found. </returns>
    private object GetMapiPropertyFromPropertyStream(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property with id '{propIdentifier}' from property stream");

        // If no property stream return null
        if (!_streamStatistics.ContainsKey(MapiTags.PropertiesStream))
            return null;

        // Get the raw bytes for the property stream
        var propBytes = GetStreamBytes(MapiTags.PropertiesStream);

        // Iterate over property stream in 16 byte chunks starting from end of header
        for (var i = _propHeaderSize; i < propBytes.Length; i = i + 16)
        {
            // Get property type located in the 1st and 2nd bytes as an unsigned short value
            var propType = (PropertyType)BitConverter.ToUInt16(propBytes, i);

            // Get property identifier located in 3rd and 4th bytes as a hexadecimal string
            var propIdent = new[] { propBytes[i + 3], propBytes[i + 2] };
            var propIdentString = BitConverter.ToString(propIdent).Replace("-", string.Empty);

            // If this is not the property being gotten continu to next property
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
                    return DateTimeOffset.FromFileTime(fileTime);

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
    ///     Gets the value of the MAPI property as a <see cref="UnsendableRecipients" />
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a string. </returns>
    private UnsendableRecipients GetUnsendableRecipients(string propIdentifier)
    {
        Logger.WriteToLog($"Getting unsendable recipients with property id '{propIdentifier}'");
        var data = GetMapiPropertyBytes(propIdentifier);
        return data != null ? new UnsendableRecipients(data) : null;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a string.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a string. </returns>
    private string GetMapiPropertyString(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property string with id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as string;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a list of string.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a list of string. </returns>
    private ReadOnlyCollection<string> GetMapiPropertyStringList(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property string list with id '{propIdentifier}'");
        var list = GetMapiProperty(propIdentifier) as List<string>;
        return list?.AsReadOnly();
    }

    /// <summary>
    ///     Gets the value of the MAPI property as an integer.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
    /// <returns> The value of the MAPI property as an integer. </returns>
    private int? GetMapiPropertyInt32(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property Int32 id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as int?;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a double.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a double. </returns>
    private double? GetMapiPropertyDouble(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property Double id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as double?;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a datetime.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a datetime or null when not set </returns>
    private DateTimeOffset? GetMapiPropertyDateTimeOffset(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property DateTime id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as DateTimeOffset?;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a bool.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a boolean or null when not set. </returns>
    private bool? GetMapiPropertyBool(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property Bool id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as bool?;
    }

    /// <summary>
    ///     Gets the value of the MAPI property as a byte array.
    /// </summary>
    /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier.</param>
    /// <returns> The value of the MAPI property as a byte array. </returns>
    private byte[] GetMapiPropertyBytes(string propIdentifier)
    {
        Logger.WriteToLog($"Getting mapi property Bytes id '{propIdentifier}'");
        return GetMapiProperty(propIdentifier) as byte[];
    }
    #endregion

    #region IDisposable Members
    /// <summary>
    ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose()
    {
        if (_rootStorage == null) return;
        
        try
        {
            if (!_leaveStreamOpen)
                _rootStorage.Dispose();
            else
                Logger.WriteToLog("The input stream is left open because the option 'leaveStreamOpen' has been set");
        }
        catch
        {
            // Ignore
        }

        _rootStorage = null;
    }
    #endregion
}