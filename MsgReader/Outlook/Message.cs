using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using DocumentServices.Modules.Readers.MsgReader.Header;
using DocumentServices.Modules.Readers.MsgReader.Helpers;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class represent a MSG object
        /// </summary>
        public class Message : Storage
        {
            #region Public enum MessageType
            /// <summary>
            /// The message types
            /// </summary>
            public enum MessageType
            {
                /// <summary>
                /// The message is an E-mail
                /// </summary>
                Email,

                /// <summary>
                /// The message is an appointment
                /// </summary>
                Appointment,

                /// <summary>
                /// The message is a request for an appointment
                /// </summary>
                AppointmentRequest,

                /// <summary>
                /// The message is a response to an appointment
                /// </summary>
                AppointmentResponse,

                /// <summary>
                /// The message is a contact card
                /// </summary>
                Contact,

                /// <summary>
                /// The message is a task
                /// </summary>
                Task,

                /// <summary>
                /// The task request accept
                /// </summary>
                TaskRequestAccept,

                /// <summary>
                /// The message is a sticky note
                /// </summary>
                StickyNote,

                /// <summary>
                /// The message type is unknown
                /// </summary>
                Unknown
            }
            #endregion

            #region MessageImportance
            /// <summary>
            /// The importancy of the message
            /// </summary>
            public enum MessageImportance
            {
                Low = 0,
                Normal = 1,
                High = 2
            }
            #endregion

            #region Fields
            /// <summary>
            /// Contains the <see cref="MessageType"/> of this Message
            /// </summary>
            private MessageType _type = MessageType.Unknown;

            /// <summary>
            /// Containts the name of the <see cref="Storage.Message"/> file
            /// </summary>
            private string _fileName;

            /// <summary>
            /// Containts all the <see cref="Storage.Recipient"/> objects
            /// </summary>
            private readonly List<Recipient> _recipients = new List<Recipient>();

            /// <summary>
            /// Contains the date/time in UTC format when the <see cref="Storage.Message"/> object has been sent
            /// </summary>
            private DateTime? _sentOn;

            /// <summary>
            /// Contains the date/time in UTC format when the <see cref="Storage.Message"/> object has been received
            /// </summary>
            private DateTime? _receivedOn;

            /// <summary>
            /// Contains the <see cref="MessageImportance"/> of the <see cref="Storage.Message"/> object
            /// </summary>
            private MessageImportance? _importance;
            /// <summary>
            /// Contains all the <see cref="Storage.Attachment"/> objects
            /// </summary>
            private readonly List<Object> _attachments = new List<Object>();

            /// <summary>
            /// Contains the subject of the <see cref="Storage.Message"/> object
            /// </summary>
            private string _subject;

            /// <summary>
            /// Contains the <see cref="Storage.Flag"/> object
            /// </summary>
            private Flag _flag;

            /// <summary>
            /// Contains the <see cref="Storage.Task"/> object
            /// </summary>
            private Task _task;

            /// <summary>
            /// Contains the <see cref="Storage.Appointment"/> object
            /// </summary>
            private Appointment _appointment;

            /// <summary>
            /// Contains the <see cref="Storage.Contact"/> object
            /// </summary>
            private Contact _contact;
            #endregion

            #region Properties
            /// <summary>
            /// Gives the <see cref="MessageType">type</see> of this message object
            /// </summary>
            public MessageType Type
            {
                get
                {
                    if (_type != MessageType.Unknown)
                        return _type;

                    var type = GetMapiPropertyString(MapiTags.PR_MESSAGE_CLASS);

                    switch (type.ToUpperInvariant())
                    {
                        case "IPM.NOTE":
                            _type = MessageType.Email;
                            break;

                        case "IPM.APPOINTMENT":
                            _type = MessageType.Appointment;
                            break;

                        case "IPM.SCHEDULE.MEETING.REQUEST":
                            _type = MessageType.AppointmentRequest;
                            break;

                        case "IPM.SCHEDULE.MEETING.RESPONSE":
                            _type = MessageType.AppointmentResponse;
                            break;

                        case "IPM.CONTACT":
                            _type = MessageType.Contact;
                            break;

                        case "IPM.TASK":
                            _type = MessageType.Task;
                            break;

                        case "IPM.TASKREQUEST.ACCEPT":
                            _type = MessageType.TaskRequestAccept;
                            break;

                        case "IPM.STICKYNOTE":
                            _type = MessageType.StickyNote;
                            break;
                    }

                    return _type;
                }
            }

            /// <summary>
            /// Returns the filename of the message object. For message object Outlook uses the subject. It strips
            /// invalid filename characters. When there is no filename the name from <see cref=" LanguageConsts.NameLessFileName"/>
            /// will be used
            /// </summary>
            public string FileName
            {
                get
                {
                    if (_fileName != null)
                        return _fileName;

                    _fileName = GetMapiPropertyString(MapiTags.PR_SUBJECT);

                    if (string.IsNullOrEmpty(_fileName))
                        _fileName = LanguageConsts.NameLessFileName;

                    _fileName = FileManager.RemoveInvalidFileNameChars(_fileName) + ".msg";
                    return _fileName;
                }
            }

            // ReSharper disable once CSharpWarnings::CS0109
            /// <summary>
            /// Gets the display value of the contact that sent the email.
            /// </summary>
            public new Sender Sender { get; private set; }

            /// <summary>
            /// Returns the list of recipients in the message object
            /// </summary>
            public List<Recipient> Recipients
            {
                get { return _recipients; }
            }

            /// <summary>
            /// Returns the date/time in UTC format when the message object has been sent, null when not available
            /// </summary>
            public DateTime? SentOn
            {
                get
                {
                    if (_sentOn != null)
                        return _sentOn;

                    _sentOn = GetMapiPropertyDateTime(MapiTags.PR_PROVIDER_SUBMIT_TIME) ??
                                 GetMapiPropertyDateTime(MapiTags.PR_CLIENT_SUBMIT_TIME);

                    if (_sentOn == null && Headers != null)
                        _sentOn = Headers.DateSent.ToLocalTime();

                    return _sentOn;
                }
            }

            /// <summary>
            /// PR_MESSAGE_DELIVERY_TIME  is the time that the message was delivered to the store and 
            /// PR_CLIENT_SUBMIT_TIME  is the time when the message was sent by the client (Outlook) to the server.
            /// Now in this case when the Outlook is offline, it refers to the local store. Therefore when an email is sent, 
            /// it gets submitted to the local store and PR_MESSAGE_DELIVERY_TIME  gets set the that time. Once the Outlook is 
            /// online at that point the message gets submitted by the client to the server and the PR_CLIENT_SUBMIT_TIME  gets stamped. 
            /// </summary>
            public DateTime? ReceivedOn
            {
                get
                {
                    if (_receivedOn != null)
                        return _receivedOn;

                    _receivedOn = GetMapiPropertyDateTime(MapiTags.PR_MESSAGE_DELIVERY_TIME);

                    if (_receivedOn == null && Headers != null && Headers.Received != null && Headers.Received.Count > 0)
                        _receivedOn = Headers.Received[0].Date.ToLocalTime();

                    return _receivedOn;
                }
            }

            /// <summary>
            /// Returns the <see cref="MessageImportance"/> of the <see cref="Storage.Message"/> object, null when not available
            /// </summary>
            public MessageImportance? Importance
            {
                get
                {
                    if (_importance != null)
                        return _importance;

                    var importance = GetMapiPropertyInt32(MapiTags.PR_IMPORTANCE);
                    if (importance == null)
                    {
                        _importance = MessageImportance.Normal;
                        return _importance;
                    }

                    switch (importance)
                    {
                        case 0:
                            _importance = MessageImportance.Low;
                            break;

                        case 1:
                            _importance = MessageImportance.Normal;
                            break;

                        case 2:
                            _importance = MessageImportance.High;
                            break;
                    }

                    return _importance;
                }
            }

            /// <summary>
            /// Returns the <see cref="MessageImportance"/> of the <see cref="Storage.Message"/> object object as text
            /// </summary>
            public string ImportanceText
            {
                get
                {
                    switch (Importance)
                    {
                        case MessageImportance.Low:
                            return LanguageConsts.ImportanceLowText;

                        case MessageImportance.Normal:
                            return LanguageConsts.ImportanceNormalText;

                        case MessageImportance.High:
                            return LanguageConsts.ImportanceHighText;

                    }

                    return LanguageConsts.ImportanceNormalText;
                }
            }

            /// <summary>
            /// Returns a list with <see cref="Storage.Attachment"/> and/or <see cref="Storage.Message"/> 
            /// objects that are attachted to the <see cref="Storage.Message"/> object
            /// </summary>
            public List<Object> Attachments
            {
                get { return _attachments; }
            }

            /// <summary>
            /// Returns the rendering position of this <see cref="Storage.Message"/> object when it was added to another
            /// <see cref="Storage.Message"/> object and the body type was set to RTF
            /// </summary>
            public int RenderingPosition { get; private set; }

            /// <summary>
            /// Returns the subject of the <see cref="Storage.Message"/> object
            /// </summary>
            public string Subject
            {
                get
                {
                    if (_subject != null)
                        return _subject;

                    _subject = GetMapiPropertyString(MapiTags.PR_SUBJECT);
                    if (string.IsNullOrEmpty(_subject))
                        _subject = string.Empty;

                    return _subject;
                }
            }

            /// <summary>
            /// Returns the available E-mail headers. These are only filled when the message
            /// has been sent accross the internet. Returns be null when there aren't
            /// any message headers
            /// </summary>
            public MessageHeader Headers { get; private set; }

            // ReSharper disable once CSharpWarnings::CS0109
            /// <summary>
            /// Returns a <see cref="Flag"/> object when a flag has been set on the <see cref="Storage.Message"/>.
            /// Returns null when not available.
            /// </summary>
            public new Flag Flag
            {
                get
                {
                    if (_flag != null)
                        return _flag;

                    var flag = new Flag(this);

                    if (flag.Request != null)
                        _flag = flag;

                    return _flag;
                }
            }

            // ReSharper disable once CSharpWarnings::CS0109
            /// <summary>
            /// Returns an <see cref="Appointment"/> object when the <see cref="MessageType"/> is a <see cref="MessageType.Appointment"/>.
            /// Returns null when not available.
            /// </summary>
            public new Appointment Appointment
            {
                get
                {
                    if (_appointment != null)
                        return _appointment;

                    switch (Type)
                    {
                        case MessageType.AppointmentRequest:
                        case MessageType.Appointment:
                        case MessageType.AppointmentResponse:
                            break;

                        default:
                            return null;
                    }

                    _appointment = new Appointment(this);
                    return _appointment;
                }
            }

            // ReSharper disable once CSharpWarnings::CS0109
            /// <summary>
            /// Returns a <see cref="Task"/> object. This property is only available when: <br/>
            /// - The <see cref="Storage.Message.Type"/> is an <see cref="Storage.Message.MessageType.Email"/> and the <see cref="Flag"/> object is not null<br/>
            /// - The <see cref="Storage.Message.Type"/> is an <see cref="Storage.Message.MessageType.Task"/> or <see cref="Storage.Message.MessageType.TaskRequestAccept"/> <br/>
            /// </summary>
            public new Task Task
            {
                get
                {
                    if (_task != null)
                        return _task;

                    switch (_type)
                    {
                        case MessageType.Email:
                            if (Flag == null)
                                return null;
                            break;

                        case MessageType.Task:
                        case MessageType.TaskRequestAccept:
                            break;

                        default:
                            return null;
                    }

                    _task = new Task(this);
                    return _task;
                }
            }

            // ReSharper disable once CSharpWarnings::CS0109
            /// <summary>
            /// Returns an <see cref="Storage.Contact"/> object when the <see cref="MessageType"/> is a <see cref="MessageType.Contact"/>.
            /// Returns null when not available.
            /// </summary>
            public new Contact Contact
            {
                get
                {
                    if (_contact != null)
                        return _contact;

                    switch (Type)
                    {
                        case MessageType.Contact:
                            break;

                        default:
                            return null;
                    }

                    _contact = new Contact(this);
                    return _contact;
                }
            }


            /// <summary>
            /// Returns the categories that are placed in the outlook message.
            /// Only supported for outlook messages from Outlook 2007 or higher
            /// </summary>
            public ReadOnlyCollection<string> Categories
            {
                get { return GetMapiPropertyStringList(MapiTags.Keywords); }
            }

            /// <summary>
            /// Returns the body of the outlook message in plain text format.
            /// </summary>
            /// <value> The body of the outlook message in plain text format. </value>
            public string BodyText
            {
                get { return GetMapiPropertyString(MapiTags.PR_BODY); }
            }

            /// <summary>
            /// Returns the body of the outlook message in RTF format.
            /// </summary>
            /// <value> The body of the outlook message in RTF format. </value>
            public string BodyRtf
            {
                get
                {
                    //get value for the RTF compressed MAPI property
                    var rtfBytes = GetMapiPropertyBytes(MapiTags.PR_RTF_COMPRESSED);

                    //return null if no property value exists
                    if (rtfBytes == null || rtfBytes.Length == 0)
                        return null;

                    //decompress the rtf value
                    rtfBytes = RtfDecompressor.DecompressRtf(rtfBytes);

                    //encode the rtf value as an ascii string and return
                    return Encoding.ASCII.GetString(rtfBytes);
                }
            }

            /// <summary>
            /// Returns the body of the outlook message in HTML format.
            /// </summary>
            /// <value> The body of the outlook message in HTML format. </value>
            public string BodyHtml
            {
                get
                {
                    // Get value for the HTML MAPI property
                    var htmlObject = GetMapiProperty(MapiTags.PR_BODY_HTML);
                    string html = null;
                    if (htmlObject is string)
                        html = htmlObject as string;
                    else if (htmlObject is byte[])
                    {
                        // Check for a code page 
                        var codePage = GetMapiPropertyInt32(MapiTags.PR_INTERNET_CPID);
                        var htmlByteArray = htmlObject as byte[];
                        var encoder = codePage == null ? Encoding.Default : Encoding.GetEncoding((int) codePage);
                        html = encoder.GetString(htmlByteArray);
                    }

                    // When there is no HTML found
                    if (html == null)
                    {
                        // Check if we have HTML embedded into rtf
                        var bodyRtf = BodyRtf;
                        if (bodyRtf != null)
                        {
                            var rtfDomDocument = new Rtf.DomDocument();
                            rtfDomDocument.LoadRtfText(bodyRtf);
                            if (!string.IsNullOrEmpty(rtfDomDocument.HtmlContent))
                                return rtfDomDocument.HtmlContent;
                        }
                    }

                    return html;
                }
            }
            #endregion

            #region Constructors
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Message" /> class from a msg file.
            /// </summary>
            /// <param name="msgfile">The msg file to load</param>
            public Message(string msgfile) : base(msgfile) { }

            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Message" /> class from a <see cref="Stream" /> containing an IStorage.
            /// </summary>
            /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
            public Message(Stream storageStream) : base(storageStream) { }

            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Message" /> class on the specified <see> <cref>NativeMethods.IStorage</cref> </see>.
            /// </summary>
            /// <param name="storage"> The storage to create the <see cref="Storage.Message" /> on. </param>
            /// <param name="renderingPosition"></param>
            public Message(NativeMethods.IStorage storage, int renderingPosition) : base(storage)
            {
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
                RenderingPosition = renderingPosition;
            }
            #endregion

            #region GetHeaders
            /// <summary>
            /// Try's to read the E-mail transport headers. They are only there when a msg file has been
            /// sent over the internet. When a message stays inside an Exchange server there are not any headers
            /// </summary>
            private void GetHeaders()
            {
                var headersString = GetMapiPropertyString(MapiTags.PR_TRANSPORT_MESSAGE_HEADERS);
                if (!string.IsNullOrEmpty(headersString))
                    Headers = HeaderExtractor.GetHeaders(headersString);
            }
            #endregion

            #region LoadStorage
            /// <summary>
            /// Processes sub storages on the specified storage to capture attachment and recipient data.
            /// </summary>
            /// <param name="storage"> The storage to check for attachment and recipient data. </param>
            protected override void LoadStorage(NativeMethods.IStorage storage)
            {
                base.LoadStorage(storage);

                GetHeaders();
                // Sender = new Sender(new Storage(storage));
                Sender = new Sender(this);

                foreach (var storageStat in _subStorageStatistics.Values)
                {
                    // Element is a storage. get it and add its statistics object to the sub storage dictionary
                    var subStorage = storage.OpenStorage(storageStat.pwcsName, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE,
                        IntPtr.Zero, 0);

                    // Run specific load method depending on sub storage name prefix
                    if (storageStat.pwcsName.StartsWith(MapiTags.RecipStoragePrefix))
                    {
                        var recipient = new Recipient(new Storage(subStorage)); 
                        _recipients.Add(recipient);
                    }
                    else if (storageStat.pwcsName.StartsWith(MapiTags.AttachStoragePrefix))
                        LoadAttachmentStorage(subStorage);
                    else
                        Marshal.ReleaseComObject(subStorage);
                }

                // Check if there is a named substorage and if so open it and map all the named MAPI properties
                if (_subStorageStatistics.ContainsKey(MapiTags.NameIdStorage))
                {
                    var mappingValues = new List<string>();

                    // Get all the named properties from the _streamStatistics
                    foreach (var streamStatistic in _streamStatistics)
                    {
                        var name = streamStatistic.Value.pwcsName;

                        if (name.StartsWith(MapiTags.SubStgVersion1))
                        {
                            // Get the property value
                            var propIdentString = name.Substring(12, 4);

                            // Convert it to a short
                            var value = ushort.Parse(propIdentString, NumberStyles.HexNumber);

                            // Check if the value is in the named property range (8000 to FFFE (Hex))
                            if (value >= 32768 && value <= 65534)
                            {
                                // If so then add it to perform mapping later on
                                if (!mappingValues.Contains(propIdentString))
                                    mappingValues.Add(propIdentString);
                            }
                        }
                    }

                    // Check if there is also a properties stream and if so get all the named MAPI properties from it
                    if (_streamStatistics.ContainsKey(MapiTags.PropertiesStream))
                    {
                        // Get the raw bytes for the property stream
                        var propBytes = GetStreamBytes(MapiTags.PropertiesStream);

                        for (var i = _propHeaderSize; i < propBytes.Length; i = i + 16)
                        {
                            // Get property identifer located in 3nd and 4th bytes as a hexdecimal string
                            var propIdent = new[] { propBytes[i + 3], propBytes[i + 2] };
                            var propIdentString = BitConverter.ToString(propIdent).Replace("-", string.Empty);

                            // Convert it to a short
                            var value = ushort.Parse(propIdentString, NumberStyles.HexNumber);

                            // Check if the value is in the named property range (8000 to FFFE (Hex))
                            if (value >= 32768 && value <= 65534)
                            {
                                // If so then add it to perform mapping later on
                                if (!mappingValues.Contains(propIdentString))
                                    mappingValues.Add(propIdentString);
                            }
                        }
                    }

                    // Check if there is something to map
                    if (mappingValues.Count <= 0) return;
                    // Get the Named Id Storage, we need this one to perform the mapping
                    var storageStat = _subStorageStatistics[MapiTags.NameIdStorage];
                    var subStorage = storage.OpenStorage(storageStat.pwcsName, IntPtr.Zero,
                        NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0);

                    // Load the subStorage into our mapping class that does all the mapping magic
                    var mapiToOom = new MapiTagMapper(new Storage(subStorage));

                    // Get the mapped properties
                    _namedProperties = mapiToOom.GetMapping(mappingValues);

                    // Clean up the com object
                    Marshal.ReleaseComObject(subStorage);
                }
            }
            #endregion

            # region LoadAttachmentStorage
            /// <summary>
            /// Loads the attachment data out of the specified storage.
            /// </summary>
            /// <param name="storage"> The attachment storage. </param>
            private void LoadAttachmentStorage(NativeMethods.IStorage storage)
            {
                // Create attachment from attachment storage
                var attachment = new Attachment(new Storage(storage));

                // If attachment is a embeded msg handle differently than an normal attachment
                var attachMethod = attachment.GetMapiPropertyInt32(MapiTags.PR_ATTACH_METHOD);

                switch (attachMethod)
                {
                    case MapiTags.ATTACH_EMBEDDED_MSG:
                        // Create new Message and set parent and header size
                        var iStorageObject =
                            attachment.GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as NativeMethods.IStorage;
                        var subMsg = new Message(iStorageObject, attachment.RenderingPosition)
                        {
                            _parentMessage = this,
                            _propHeaderSize = MapiTags.PropertiesStreamHeaderEmbeded
                        };
                        _attachments.Add(subMsg);
                        break;

                    default:
                        // Add attachment to attachment list
                        _attachments.Add(attachment);
                        break;
                }
            }
            #endregion

            #region Save
            /// <summary>
            /// Saves this <see cref="Storage.Message" /> to the specified file name.
            /// </summary>
            /// <param name="fileName"> Name of the file. </param>
            public void Save(string fileName)
            {
                var saveFileStream = File.Open(fileName, FileMode.Create, FileAccess.ReadWrite);
                Save(saveFileStream);
                saveFileStream.Close();
            }

            /// <summary>
            /// Saves this <see cref="Storage.Message"/> to the specified stream.
            /// </summary>
            /// <param name="stream"> The stream to save to. </param>
            public void Save(Stream stream)
            {
                // Get statistics for stream 
                Storage saveMsg = this;

                NativeMethods.IStorage memoryStorage = null;
                NativeMethods.IStorage nameIdSourceStorage = null;
                NativeMethods.ILockBytes memoryStorageBytes = null;

                try
                {
                    // Create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    NativeMethods.StgCreateDocfileOnILockBytes(memoryStorageBytes, NativeMethods.STGM.CREATE | NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, out memoryStorage);

                    // Copy the save storage into the new storage
                    saveMsg._storage.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    // If not the top parent then the name id mapping needs to be copied from top parent to this message and the property stream header needs to be padded by 8 bytes
                    if (!IsTopParent)
                    {
                        // Create a new name id storage and get the source name id storage to copy from
                        var nameIdStorage = memoryStorage.CreateStorage(MapiTags.NameIdStorage, NativeMethods.STGM.CREATE | NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, 0);
                        nameIdSourceStorage = TopParent._storage.OpenStorage(MapiTags.NameIdStorage, IntPtr.Zero, NativeMethods.STGM.READ | NativeMethods.STGM.SHARE_EXCLUSIVE,
                            IntPtr.Zero, 0);

                        // Copy the name id storage from the parent to the new name id storage
                        nameIdSourceStorage.CopyTo(0, null, IntPtr.Zero, nameIdStorage);

                        // Get the property bytes for the storage being copied
                        var props = saveMsg.GetStreamBytes(MapiTags.PropertiesStream);

                        // Create new array to store a copy of the properties that is 8 bytes larger than the old so the header can be padded
                        var newProps = new byte[props.Length + 8];

                        // Insert 8 null bytes from index 24 to 32. this is because a top level object property header requires a 32 byte header
                        Buffer.BlockCopy(props, 0, newProps, 0, 24);
                        Buffer.BlockCopy(props, 24, newProps, 32, props.Length - 24);

                        // Remove the copied prop bytes so it can be replaced with the padded version
                        memoryStorage.DestroyElement(MapiTags.PropertiesStream);

                        // Create the property stream again and write in the padded version
                        var propStream = memoryStorage.CreateStream(MapiTags.PropertiesStream, NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, 0, 0);
                        propStream.Write(newProps, newProps.Length, IntPtr.Zero);
                    }

                    // Commit changes to the storage
                    memoryStorage.Commit(0);
                    memoryStorageBytes.Flush();

                    // Get the STATSTG of the ILockBytes to determine how many bytes were written to it
                    STATSTG memoryStorageBytesStat;
                    memoryStorageBytes.Stat(out memoryStorageBytesStat, 1);

                    // Read the bytes into a managed byte array
                    var memoryStorageContent = new byte[memoryStorageBytesStat.cbSize];
                    memoryStorageBytes.ReadAt(0, memoryStorageContent, memoryStorageContent.Length, null);

                    // Write storage bytes to stream
                    stream.Write(memoryStorageContent, 0, memoryStorageContent.Length);
                }
                finally
                {
                    if (nameIdSourceStorage != null)
                        Marshal.ReleaseComObject(nameIdSourceStorage);

                    if (memoryStorage != null)
                        Marshal.ReleaseComObject(memoryStorage);

                    if (memoryStorageBytes != null)
                        Marshal.ReleaseComObject(memoryStorageBytes);
                }
            }
            #endregion

            #region Disposing
            /// <summary>
            /// Dispose of all the objects
            /// </summary>
            protected override void Disposing()
            {
                // Dispose sub storages
                foreach (var recipient in _recipients)
                    recipient.Dispose();

                // Dispose sub storages
                foreach (var attachment in _attachments)
                {
                    var tempAttachment = attachment as Attachment;
                    if (tempAttachment != null)
                        tempAttachment.Dispose();
                    else
                    {
                        var message = attachment as Message;
                        if (message != null)
                            message.Dispose();
                    }
                }
            }
            #endregion
        }
    }
}