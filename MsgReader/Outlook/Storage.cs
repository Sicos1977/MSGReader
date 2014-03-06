using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
//using System.Windows.Forms;
using System.Web;
using DocumentServices.Modules.Readers.MsgReader.Header;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    // ReSharper disable LocalizableElement
    // ReSharper disable DoNotCallOverridableMethodsInConstructor
    public class Storage : IDisposable
    {
        #region Class NativeMethods
        protected static class NativeMethods
        {
            #region Stgm enum
            [Flags]
            public enum Stgm
            {
                Direct = 0x00000000,
                Transacted = 0x00010000,
                Simple = 0x08000000,
                Read = 0x00000000,
                Write = 0x00000001,
                Readwrite = 0x00000002,
                ShareDenyNone = 0x00000040,
                ShareDenyRead = 0x00000030,
                ShareDenyWrite = 0x00000020,
                ShareExclusive = 0x00000010,
                Priority = 0x00040000,
                Deleteonrelease = 0x04000000,
                Noscratch = 0x00100000,
                Create = 0x00001000,
                Convert = 0x00020000,
                Failifthere = 0x00000000,
                Nosnapshot = 0x00200000,
                DirectSwmr = 0x00400000
            }
            #endregion

            #region Consts
            public const ushort PtUnspecified = 0; /* (Reserved for interface use) type doesn't matter to caller */
            public const ushort PtNull = 1; /* NULL property value */
            public const ushort PtI2 = 2; /* Signed 16-bit value */
            public const ushort PtLong = 3; /* Signed 32-bit value */
            public const ushort PtR4 = 4; /* 4-byte floating point */
            public const ushort PtDouble = 5; /* Floating point double */
            public const ushort PtCurrency = 6; /* Signed 64-bit int (decimal w/    4 digits right of decimal pt) */
            public const ushort PtApptime = 7; /* Application time */
            public const ushort PtError = 10; /* 32-bit error value */
            public const ushort PtBoolean = 11; /* 16-bit boolean (non-zero true) */
            public const ushort PtObject = 13; /* Embedded object in a property */
            public const ushort PtI8 = 20; /* 8-byte signed integer */
            public const ushort PtString8 = 30; /* Null terminated 8-bit character string */
            public const ushort PtUnicode = 31; /* Null terminated Unicode string */
            public const ushort PtSystime = 64; /* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
            public const ushort PtClsid = 72; /* OLE GUID */
            public const ushort PtBinary = 258; /* Uninterpreted (counted byte array) */
            #endregion

            #region DllImports
            [DllImport("ole32.DLL")]
            public static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease,
                out ILockBytes ppLkbyt);

            [DllImport("ole32.DLL")]
            public static extern int StgIsStorageILockBytes(ILockBytes plkbyt);

            [DllImport("ole32.DLL")]
            public static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, Stgm grfMode, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            public static extern void StgOpenStorageOnILockBytes(ILockBytes plkbyt, IStorage pstgPriority, Stgm grfMode,
                IntPtr snbExclude, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            public static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string wcsName);

            [DllImport("ole32.DLL")]
            public static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string wcsName,
                IStorage pstgPriority, Stgm grfMode, IntPtr snbExclude, int reserved,
                out IStorage ppstgOpen);
            #endregion

            #region CloneStorage
            public static IStorage CloneStorage(IStorage source, bool closeSource)
            {
                IStorage memoryStorage = null;
                ILockBytes memoryStorageBytes = null;
                try
                {
                    //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    StgCreateDocfileOnILockBytes(memoryStorageBytes, Stgm.Create | Stgm.Readwrite | Stgm.ShareExclusive,
                        0, out memoryStorage);

                    //copy the source storage into the new storage
                    source.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    //ensure memory is released
                    ReferenceManager.AddItem(memoryStorage);
                }
                catch
                {
                    if (memoryStorage != null)
                        Marshal.ReleaseComObject(memoryStorage);
                }
                finally
                {
                    if (memoryStorageBytes != null)
                        Marshal.ReleaseComObject(memoryStorageBytes);

                    if (closeSource)
                        Marshal.ReleaseComObject(source);
                }

                return memoryStorage;
            }
            #endregion

            #region Nested type: IEnumSTATSTG
            [ComImport, Guid("0000000D-0000-0000-C000-000000000046"),
             InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IEnumSTATSTG
            {
                void Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] STATSTG[] rgelt, out uint pceltFetched);
                void Skip(uint celt);
                void Reset();

                [return: MarshalAs(UnmanagedType.Interface)]
                IEnumSTATSTG Clone();
            }
            #endregion

            #region Nested type: ILockBytes
            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"),
             InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface ILockBytes
            {
                void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset,
                    [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                    [In, MarshalAs(UnmanagedType.U4)] int cb,
                    [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);

                void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset,
                    [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                    [In, MarshalAs(UnmanagedType.U4)] int cb,
                    [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);

                void Flush();
                void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);

                void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset,
                    [In, MarshalAs(UnmanagedType.U8)] long cb,
                    [In, MarshalAs(UnmanagedType.U4)] int dwLockType);

                void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset,
                    [In, MarshalAs(UnmanagedType.U8)] long cb,
                    [In, MarshalAs(UnmanagedType.U4)] int dwLockType);

                void Stat([Out] out STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
            }
            #endregion

            #region Nested type: IStorage
            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
             Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode, IntPtr snbExclude,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved);

                void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIidExclude,
                    IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);

                void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest,
                    [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName,
                    [In, MarshalAs(UnmanagedType.U4)] int grfFlags);

                void Commit(int grfCommitFlags);
                void Revert();

                void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved3,
                    [MarshalAs(UnmanagedType.Interface)] out IEnumSTATSTG ppVal);

                void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);

                void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName,
                    [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);

                void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] FILETIME pctime,
                    [In] FILETIME patime, [In] FILETIME pmtime);

                void SetClass([In] ref Guid clsid);
                void SetStateBits(int grfStateBits, int grfMask);
                void Stat([Out] out STATSTG pStatStg, int grfStatFlag);
            }
            #endregion
        }
        #endregion

        #region Class ReferenceManager
        private class ReferenceManager
        {
            private static readonly ReferenceManager Instance = new ReferenceManager();

            private readonly List<object> _trackingObjects = new List<object>();

            private ReferenceManager()
            {
            }

            public static void AddItem(object track)
            {
                lock (Instance)
                {
                    if (!Instance._trackingObjects.Contains(track))
                        Instance._trackingObjects.Add(track);
                }
            }

            public static void RemoveItem(object track)
            {
                lock (Instance)
                {
                    if (Instance._trackingObjects.Contains(track))
                        Instance._trackingObjects.Remove(track);
                }
            }

            ~ReferenceManager()
            {
                foreach (var trackingObject in _trackingObjects)
                    Marshal.ReleaseComObject(trackingObject);
            }
        }
        #endregion

        #region Enum RecipientType
        public enum RecipientType
        {
            To,
            Cc,
            Unknown
        }
        #endregion

        #region Nested class Attachment
        public class Attachment : Storage
        {
            #region Property(s)
            /// <summary>
            /// Gets the filename.
            /// </summary>
            /// <value> The filename. </value>
            public string Filename
            {
                get
                {
                    var filename = GetMapiPropertyString(PrAttachLongFilename);
                    if (string.IsNullOrEmpty(filename))
                    {
                        filename = GetMapiPropertyString(PrAttachFilename);
                    }
                    if (string.IsNullOrEmpty(filename))
                    {
                        filename = GetMapiPropertyString(PrDisplayName);
                    }
                    return filename;
                }
            }

            /// <summary>
            /// Gets the data.
            /// </summary>
            /// <value> The data. </value>
            public byte[] Data
            {
                get { return GetMapiPropertyBytes(PrAttachData); }
            }

            /// <summary>
            /// Gets the content id.
            /// </summary>
            /// <value> The content id. </value>
            public string ContentId
            {
                get { return GetMapiPropertyString(PrAttachContentId); }
            }

            /// <summary>
            /// Gets the rendering posisiton.
            /// </summary>
            /// <value> The rendering posisiton. </value>
            public int RenderingPosisiton
            {
                get { return GetMapiPropertyInt32(PrRenderingPosition); }
            }
            #endregion

            #region Constructor(s)
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Attachment" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            public Attachment(Storage message)
                : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = PropertiesStreamHeaderAttachOrRecip;
            }
            #endregion
        }
        #endregion

        #region Nested class Header
        public class Header
        {
            #region Properties
            /// <summary>
            /// The name of the header value
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// The value of the header
            /// </summary>
            public string Value { get; set; }
            #endregion
        }
        #endregion

        #region Nested class Message
        public class Message : Storage
        {
            #region Properties
            /// <summary>
            /// Containts any attachments
            /// </summary>
            private readonly List<Object> _attachments = new List<Object>();

            /// <summary>
            /// Containts any MSG attachments
            /// </summary>
            //private readonly List<Message> _messages = new List<Message>();

            /// <summary>
            /// Containts all the recipients
            /// </summary>
            private readonly List<Recipient> _recipients = new List<Recipient>();

            /// <summary>
            /// Gets the list of recipients in the outlook message.
            /// </summary>
            /// <value>The list of recipients in the outlook message</value>
            public List<Recipient> Recipients
            {
                get { return _recipients; }
            }

            /// <summary>
            /// Gets the list of attachments in the outlook message.
            /// </summary>
            /// <value>The list of attachments in the outlook message</value>
            public List<Object> Attachments
            {
                get { return _attachments; }
            }

            ///// <summary>
            ///// Gets the list of sub messages in the outlook message.
            ///// </summary>
            ///// <value>The list of sub messages in the outlook message</value>
            //public List<Message> Messages
            //{
            //    get { return _messages; }
            //}

            /// <summary>
            /// Gets the display value of the contact that sent the email.
            /// </summary>
            /// <value>The display value of the contact that sent the email</value>
            public Sender Sender { get; private set; }

            /// <summary>
            /// Gives the aviable E-mail headers. These are only filled when the message
            /// has been sent accross the internet. This will be null when there aren't
            /// any message headers
            /// </summary>
            public MessageHeader Headers { get; private set; }

            /// <summary>
            /// Gets the date/time in UTC format when the message is sent.
            /// Null when not available
            /// </summary>
            public DateTime? SentOn
            {
                get
                {
                    var sentOn = GetMapiPropertyString(PrClientSubmitTime);

                    if (sentOn != null)
                    {
                        DateTime dateTime;
                        if (DateTime.TryParse(sentOn, out dateTime))
                            return dateTime;
                    }

                    if (Headers != null)
                        return Headers.DateSent;

                    return null;
                }
            }

            public DateTime? ReceivedOn
            {
                get
                {
                    var receivedOn = GetMapiPropertyString(PrMessageDeliveryTime);

                    if (receivedOn != null)
                    {
                        DateTime dateTime;
                        if (DateTime.TryParse(receivedOn, out dateTime))
                            return dateTime;
                    }

                    return null;
                }
            }

            /// <summary>
            /// Gets the subject of the outlook message.
            /// </summary>
            /// <value> The subject of the outlook message. </value>
            public string Subject
            {
                get { return GetMapiPropertyString(PrSubject); }
            }

            /// <summary>
            /// Gets the body of the outlook message in plain text format.
            /// </summary>
            /// <value> The body of the outlook message in plain text format. </value>
            public string BodyText
            {
                get { return GetMapiPropertyString(PrBody); }
            }

            /// <summary>
            /// Gets the body of the outlook message in RTF format.
            /// </summary>
            /// <value> The body of the outlook message in RTF format. </value>
            public string BodyRtf
            {
                get
                {
                    //get value for the RTF compressed MAPI property
                    var rtfBytes = GetMapiPropertyBytes(PrRtfCompressed);

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
            /// Gets the body of the outlook message in HTML format.
            /// </summary>
            /// <value> The body of the outlook message in HTML format. </value>
            public string BodyHtml
            {
                get
                {
                    //get value for the HTML MAPI property
                    var html = GetMapiPropertyString(PrBodyHtml);

                    // Als er geen HTML gedeelte is gevonden
                    if (html == null)
                    {
                        // Check if we have html embedded into rtf
                        var bodyRtf = BodyRtf;
                        if (bodyRtf != null)
                        {
                            var rtfDomDocument = new Rtf.DomDocument();
                            rtfDomDocument.LoadRtfText(bodyRtf);
                            if (!string.IsNullOrEmpty(rtfDomDocument.HtmlContent))
                                return rtfDomDocument.HtmlContent;
                        }

                        return null;
                    }

                    return html;
                }
            }
            #endregion

            #region Constructor(s)
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Message" /> class from a msg file.
            /// </summary>
            /// <param name="msgfile">The msg file to load</param>
            public Message(string msgfile)
                : base(msgfile)
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Message" /> class from a <see cref="Stream" /> containing an IStorage.
            /// </summary>
            /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
            public Message(Stream storageStream)
                : base(storageStream)
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Message" /> class on the specified <see> <cref>NativeMethods.IStorage</cref> </see> .
            /// </summary>
            /// <param name="storage"> The storage to create the <see cref="Storage.Message" /> on. </param>
            private Message(NativeMethods.IStorage storage)
                : base(storage)
            {
                _propHeaderSize = PropertiesStreamHeaderTop;
            }
            #endregion

            #region GetHeaders()
            /// <summary>
            /// Try's to read the E-mail transport headers. They are only there when a msg file has been
            /// sent over the internet. When a message stays inside an exchange server there are not any headers
            /// </summary>
            private void GetHeaders()
            {
                // According to Microsoft the headers should be in PrTransportMessageHeaders1
                // but in my case they are always in PrTransportMessageHeaders2 ... meaby that this
                // has something to do with that I use Outlook 2010??
                var headersString = GetMapiPropertyString(PrTransportMessageHeaders1);
                if (string.IsNullOrEmpty(headersString))
                    headersString = GetMapiPropertyString(PrTransportMessageHeaders2);

                if (!string.IsNullOrEmpty(headersString))
                    Headers = HeaderExtractor.GetHeaders(headersString);
            }
            #endregion

            #region Methods(LoadStorage)
            /// <summary>
            /// Processes sub storages on the specified storage to capture attachment and recipient data.
            /// </summary>
            /// <param name="storage"> The storage to check for attachment and recipient data. </param>
            protected override void LoadStorage(NativeMethods.IStorage storage)
            {
                base.LoadStorage(storage);
                Sender = new Sender(new Storage(storage));
                GetHeaders();

                foreach (var storageStat in SubStorageStatistics.Values)
                {
                    //element is a storage. get it and add its statistics object to the sub storage dictionary
                    var subStorage = storage.OpenStorage(storageStat.pwcsName, IntPtr.Zero,
                        NativeMethods.Stgm.Read |
                        NativeMethods.Stgm.ShareExclusive,
                        IntPtr.Zero, 0);


                    //run specific load method depending on sub storage name prefix
                    if (storageStat.pwcsName.StartsWith(RecipStoragePrefix))
                    {
                        var recipient = new Recipient(new Storage(subStorage));
                        _recipients.Add(recipient);
                    }
                    else if (storageStat.pwcsName.StartsWith(AttachStoragePrefix))
                    {
                        LoadAttachmentStorage(subStorage);
                    }
                    else
                    {
                        //release sub storage
                        Marshal.ReleaseComObject(subStorage);
                    }
                }
            }

            /// <summary>
            /// Loads the attachment data out of the specified storage.
            /// </summary>
            /// <param name="storage"> The attachment storage. </param>
            private void LoadAttachmentStorage(NativeMethods.IStorage storage)
            {
                //create attachment from attachment storage
                var attachment = new Attachment(new Storage(storage));

                //if attachment is a embeded msg handle differently than an normal attachment
                var attachMethod = attachment.GetMapiPropertyInt32(PrAttachMethod);
                if (attachMethod == AttachEmbeddedMsg)
                {
                    //create new Message and set parent and header size
                    var subMsg = new Message(attachment.GetMapiProperty(PrAttachData) as NativeMethods.IStorage) { _parentMessage = this, _propHeaderSize = PropertiesStreamHeaderEmbeded };
                    _attachments.Add(subMsg);
                    //add to messages list
                    //_messages.Add(subMsg);
                }
                else
                {
                    //add attachment to attachment list
                    _attachments.Add(attachment);
                }
            }
            #endregion

            #region Methods(Save)
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
            /// Saves this <see cref="Storage.Message" /> to the specified stream.
            /// </summary>
            /// <param name="stream"> The stream to save to. </param>
            public void Save(Stream stream)
            {
                //get statistics for stream 
                Storage saveMsg = this;

                NativeMethods.IStorage memoryStorage = null;
                NativeMethods.IStorage nameIdSourceStorage = null;
                NativeMethods.ILockBytes memoryStorageBytes = null;
                try
                {
                    //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    NativeMethods.StgCreateDocfileOnILockBytes(memoryStorageBytes,
                        NativeMethods.Stgm.Create | NativeMethods.Stgm.Readwrite |
                        NativeMethods.Stgm.ShareExclusive, 0, out memoryStorage);

                    //copy the save storage into the new storage
                    saveMsg._storage.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    //if not the top parent then the name id mapping needs to be copied from top parent to this message and the property stream header needs to be padded by 8 bytes
                    if (!IsTopParent)
                    {
                        //create a new name id storage and get the source name id storage to copy from
                        var nameIdStorage = memoryStorage.CreateStorage(NameidStorage,
                            NativeMethods.Stgm.Create |
                            NativeMethods.Stgm.Readwrite |
                            NativeMethods.Stgm.ShareExclusive, 0, 0);
                        nameIdSourceStorage = TopParent._storage.OpenStorage(NameidStorage, IntPtr.Zero,
                            NativeMethods.Stgm.Read |
                            NativeMethods.Stgm.ShareExclusive,
                            IntPtr.Zero, 0);

                        //copy the name id storage from the parent to the new name id storage
                        nameIdSourceStorage.CopyTo(0, null, IntPtr.Zero, nameIdStorage);

                        //get the property bytes for the storage being copied
                        var props = saveMsg.GetStreamBytes(PropertiesStream);

                        //create new array to store a copy of the properties that is 8 bytes larger than the old so the header can be padded
                        var newProps = new byte[props.Length + 8];

                        //insert 8 null bytes from index 24 to 32. this is because a top level object property header requires a 32 byte header
                        Buffer.BlockCopy(props, 0, newProps, 0, 24);
                        Buffer.BlockCopy(props, 24, newProps, 32, props.Length - 24);

                        //remove the copied prop bytes so it can be replaced with the padded version
                        memoryStorage.DestroyElement(PropertiesStream);

                        //create the property stream again and write in the padded version
                        var propStream = memoryStorage.CreateStream(PropertiesStream,
                            NativeMethods.Stgm.Readwrite |
                            NativeMethods.Stgm.ShareExclusive, 0, 0);
                        propStream.Write(newProps, newProps.Length, IntPtr.Zero);
                    }

                    //commit changes to the storage
                    memoryStorage.Commit(0);
                    memoryStorageBytes.Flush();

                    //get the STATSTG of the ILockBytes to determine how many bytes were written to it
                    STATSTG memoryStorageBytesStat;
                    memoryStorageBytes.Stat(out memoryStorageBytesStat, 1);

                    //read the bytes into a managed byte array
                    var memoryStorageContent = new byte[memoryStorageBytesStat.cbSize];
                    memoryStorageBytes.ReadAt(0, memoryStorageContent, memoryStorageContent.Length, null);

                    //write storage bytes to stream
                    stream.Write(memoryStorageContent, 0, memoryStorageContent.Length);
                }
                finally
                {
                    if (nameIdSourceStorage != null)
                    {
                        Marshal.ReleaseComObject(nameIdSourceStorage);
                    }

                    if (memoryStorage != null)
                    {
                        Marshal.ReleaseComObject(memoryStorage);
                    }

                    if (memoryStorageBytes != null)
                    {
                        Marshal.ReleaseComObject(memoryStorageBytes);
                    }
                }
            }
            #endregion

            #region Methods(Disposing)
            protected override void Disposing()
            {
                //dispose sub storages
                //foreach (var subMsg in _messages)
                //{
                //    subMsg.Dispose();
                //}

                //dispose sub storages
                foreach (var recip in _recipients)
                {
                    recip.Dispose();
                }

                //dispose sub storages
                foreach (var attach in _attachments)
                {
                    if (attach.GetType() == typeof (Attachment))
                        ((Attachment) attach).Dispose();

                    if (attach.GetType() == typeof(Message))
                        ((Message)attach).Dispose();
                }
            }
            #endregion
        }
        #endregion

        #region Nested class Sender
        public class Sender : Storage
        {
            /// <summary>
            /// Gets the display value of the contact that sent the email.
            /// </summary>
            public string DisplayName
            {
                get { return GetMapiPropertyString(PrSenderName); }
            }

            /// <summary>
            /// Gets the sender email
            /// </summary>
            public string Email
            {
                get
                {
                    var eMail = GetMapiPropertyString(PrSenderEmail);

                    if (string.IsNullOrEmpty(eMail) || eMail.IndexOf('@') < 0)
                    {
                        try
                        {
                            eMail = GetMapiPropertyString(PrSenderEmail2);
                        }
                        // ReSharper disable EmptyGeneralCatchClause
                        catch
                        {
                        }
                        // ReSharper restore EmptyGeneralCatchClause
                    }

                    if (string.IsNullOrEmpty(eMail) || eMail.IndexOf("@", StringComparison.Ordinal) < 0)
                    {
                        // get address from email header
                        var header = GetStreamAsString(HeaderStreamName, Encoding.Unicode);
                        var m = Regex.Match(header, "From:.*<(?<email>.*?)>");
                        eMail = m.Groups["email"].ToString();
                    }

                    return eMail;
                }
            }

            #region Constructor(s)
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Sender" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            public Sender(Storage message)
                : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = PropertiesStreamHeaderAttachOrRecip;
            }
            #endregion
        }
        #endregion

        #region Nested class Recipient
        public class Recipient : Storage
        {
            #region Property(s)
            /// <summary>
            /// Gets the display name.
            /// </summary>
            /// <value> The display name. </value>
            public string DisplayName
            {
                get { return GetMapiPropertyString(PrDisplayName); }
            }

            /// <summary>
            /// Gets the recipient email.
            /// </summary>
            /// <value> The recipient email. </value>
            public string Email
            {
                get
                {
                    var email = GetMapiPropertyString(PrEmail);

                    if (string.IsNullOrEmpty(email))
                        email = GetMapiPropertyString(PrEmail2);

                    return email;
                }
            }

            /// <summary>
            /// Gets the recipient type.
            /// </summary>
            /// <value> The recipient type. </value>
            public RecipientType Type
            {
                get
                {
                    var recipientType = GetMapiPropertyInt32(PrRecipientType);
                    switch (recipientType)
                    {
                        case MapiTo:
                            return RecipientType.To;

                        case MapiCc:
                            return RecipientType.Cc;
                    }
                    return RecipientType.Unknown;
                }
            }
            #endregion

            #region Constructor(s)
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Recipient" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            public Recipient(Storage message)
                : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = PropertiesStreamHeaderAttachOrRecip;
            }
            #endregion
        }
        #endregion

        #region Constants
        // Attachment constants
        private const string AttachStoragePrefix = "__attach_version1.0_#";
        private const string PrAttachFilename = "3704";
        private const string PrAttachLongFilename = "3707";
        private const string PrAttachData = "3701";
        private const string PrAttachMethod = "3705";
        private const string PrRenderingPosition = "370B";
        private const string PrAttachContentId = "3712";
        private const int AttachEmbeddedMsg = 5;

        // Recipient constants
        private const string RecipStoragePrefix = "__recip_version1.0_#";
        private const string PrDisplayName = "3001";
        private const string PrEmail = "39FE";
        private const string PrEmail2 = "403E";

        private const string PrRecipientType = "0C15";
        private const int MapiTo = 1;
        private const int MapiCc = 2;

        // MSG constants
        private const string PrSubject = "0037";
        private const string PrClientSubmitTime = "0039";
        private const string PrMessageDeliveryTime = "0E06";

        //private const string PrCreationTime = "3007";
        //private const string PrLastModificationTime = "3008";
        private const string PrBody = "1000";
        private const string PrBodyHtml = "1013";
        private const string PrRtfCompressed = "1009";
        // ReSharper disable once UnusedMember.Local
        private const string PrInternetCpid = "3FDE"; //Gives the codepage used in the body or html
        private const string PrSenderName = "0C1A";
        private const string PrSenderEmail = "0C1F";
        private const string PrSenderEmail2 = "8012";
        private const string PrTransportMessageHeaders1 = "007D001E";
        private const string PrTransportMessageHeaders2 = "007D001F";

        private const string HeaderStreamName = "__substg1.0_007D001F";
        //property stream constants
        private const string PropertiesStream = "__properties_version1.0";
        private const int PropertiesStreamHeaderTop = 32;
        private const int PropertiesStreamHeaderEmbeded = 24;
        private const int PropertiesStreamHeaderAttachOrRecip = 8;

        // Name id storage name in root storage
        private const string NameidStorage = "__nameid_version1.0";
        #endregion

        #region Properties
        /// <summary>
        /// The statistics for all streams in the IStorage associated with this instance.
        /// </summary>
        public Dictionary<string, STATSTG> StreamStatistics = new Dictionary<string, STATSTG>();

        /// <summary>
        /// The statistics for all storgages in the IStorage associated with this instance.
        /// </summary>
        public Dictionary<string, STATSTG> SubStorageStatistics = new Dictionary<string, STATSTG>();

        /// <summary>
        /// Indicates wether this instance has been disposed.
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// A reference to the parent message that this message may belong to.
        /// </summary>
        private Storage _parentMessage;

        /// <summary>
        /// Header size of the property stream in the IStorage associated with this instance.
        /// </summary>
        private int _propHeaderSize = PropertiesStreamHeaderTop;

        /// <summary>
        /// The IStorage associated with this instance.
        /// </summary>
        private NativeMethods.IStorage _storage;

        /// <summary>
        /// Gets the top level outlook message from a sub message at any level.
        /// </summary>
        /// <value> The top level outlook message. </value>
        private Storage TopParent
        {
            get { return _parentMessage != null ? _parentMessage.TopParent : this; }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is the top level outlook message.
        /// </summary>
        /// <value> <c>true</c> if this instance is the top level outlook message; otherwise, <c>false</c> . </value>
        private bool IsTopParent
        {
            get { return _parentMessage == null; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class from a file.
        /// </summary>
        /// <param name="storageFilePath"> The file to load. </param>
        private Storage(string storageFilePath)
        {
            //ensure provided file is an IStorage
            if (NativeMethods.StgIsStorageFile(storageFilePath) != 0)
            {
                throw new ArgumentException("The provided file is not a valid IStorage", "storageFilePath");
            }

            //open and load IStorage from file
            NativeMethods.IStorage fileStorage;
            NativeMethods.StgOpenStorage(storageFilePath, null,
                NativeMethods.Stgm.Read | NativeMethods.Stgm.ShareDenyWrite, IntPtr.Zero, 0,
                out fileStorage);
            LoadStorage(fileStorage);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class from a <see cref="Stream" /> containing an IStorage.
        /// </summary>
        /// <param name="storageStream"> The <see cref="Stream" /> containing an IStorage. </param>
        private Storage(Stream storageStream)
        {
            NativeMethods.IStorage memoryStorage = null;
            NativeMethods.ILockBytes memoryStorageBytes = null;
            try
            {
                //read stream into buffer
                var buffer = new byte[storageStream.Length];
                storageStream.Read(buffer, 0, buffer.Length);

                //create a ILockBytes (unmanaged byte array) and write buffer into it
                NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                memoryStorageBytes.WriteAt(0, buffer, buffer.Length, null);

                //ensure provided stream data is an IStorage
                if (NativeMethods.StgIsStorageILockBytes(memoryStorageBytes) != 0)
                {
                    throw new ArgumentException("The provided stream is not a valid IStorage", "storageStream");
                }

                //open and load IStorage on the ILockBytes
                NativeMethods.StgOpenStorageOnILockBytes(memoryStorageBytes, null,
                    NativeMethods.Stgm.Read | NativeMethods.Stgm.ShareDenyWrite,
                    IntPtr.Zero, 0, out memoryStorage);
                LoadStorage(memoryStorage);
            }
            catch
            {
                if (memoryStorage != null)
                {
                    Marshal.ReleaseComObject(memoryStorage);
                }
            }
            finally
            {
                if (memoryStorageBytes != null)
                {
                    Marshal.ReleaseComObject(memoryStorageBytes);
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Storage" /> class on the specified <see cref="NativeMethods.IStorage" />.
        /// </summary>
        /// <param name="storage"> The storage to create the <see cref="Storage" /> on. </param>
        private Storage(NativeMethods.IStorage storage)
        {
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
        /// Processes sub streams and storages on the specified storage.
        /// </summary>
        /// <param name="storage"> The storage to get sub streams and storages for. </param>
        protected virtual void LoadStorage(NativeMethods.IStorage storage)
        {
            _storage = storage;

            //ensures memory is released
            ReferenceManager.AddItem(storage);

            NativeMethods.IEnumSTATSTG storageElementEnum = null;
            try
            {
                //enum all elements of the storage
                storage.EnumElements(0, IntPtr.Zero, 0, out storageElementEnum);

                //iterate elements
                while (true)
                {
                    //get 1 element out of the com enumerator
                    uint elementStatCount;
                    var elementStats = new STATSTG[1];
                    storageElementEnum.Next(1, elementStats, out elementStatCount);

                    //break loop if element not retrieved
                    if (elementStatCount != 1)
                    {
                        break;
                    }

                    var elementStat = elementStats[0];
                    switch (elementStat.type)
                    {
                        case 1:
                            //element is a storage. add its statistics object to the storage dictionary
                            SubStorageStatistics.Add(elementStat.pwcsName, elementStat);
                            break;

                        case 2:
                            //element is a stream. add its statistics object to the stream dictionary
                            StreamStatistics.Add(elementStat.pwcsName, elementStat);
                            break;
                    }
                }
            }
            finally
            {
                //free memory
                if (storageElementEnum != null)
                {
                    Marshal.ReleaseComObject(storageElementEnum);
                }
            }
        }
        #endregion

        #region GetStreamBytes
        /// <summary>
        /// Gets the data in the specified stream as a byte array.
        /// </summary>
        /// <param name="streamName"> Name of the stream to get data for. </param>
        /// <returns> A byte array containg the stream data. </returns>
        public byte[] GetStreamBytes(string streamName)
        {
            //get statistics for stream 
            var streamStatStg = StreamStatistics[streamName];

            byte[] iStreamContent;
            IStream stream = null;
            try
            {
                //open stream from the storage
                stream = _storage.OpenStream(streamStatStg.pwcsName, IntPtr.Zero,
                    NativeMethods.Stgm.Read | NativeMethods.Stgm.ShareExclusive, 0);

                //read the stream into a managed byte array
                iStreamContent = new byte[streamStatStg.cbSize];
                stream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);
            }
            finally
            {
                if (stream != null)
                {
                    Marshal.ReleaseComObject(stream);
                }
            }

            //return the stream bytes
            return iStreamContent;
        }
        #endregion

        #region GetStreamAsString
        /// <summary>
        /// Gets the data in the specified stream as a string using the specifed encoding to decode the stream data.
        /// </summary>
        /// <param name="streamName"> Name of the stream to get string data for. </param>
        /// <param name="streamEncoding"> The encoding to decode the stream data with. </param>
        /// <returns> The data in the specified stream as a string. </returns>
        public string GetStreamAsString(string streamName, Encoding streamEncoding)
        {
            try
            {
                var streamReader = new StreamReader(new MemoryStream(GetStreamBytes(streamName)), streamEncoding);
                var streamContent = streamReader.ReadToEnd();
                streamReader.Close();

                return streamContent;
            }
            catch (Exception)
            {
                return string.Empty;
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
            //try get prop value from stream or storage
            var propValue = GetMapiPropertyFromStreamOrStorage(propIdentifier) ??
                            GetMapiPropertyFromPropertyStream(propIdentifier);

            //if not found in stream or storage try get prop value from property stream
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
            //get list of stream and storage identifiers which map to properties
            var propKeys = new List<string>();
            propKeys.AddRange(StreamStatistics.Keys);
            propKeys.AddRange(SubStorageStatistics.Keys);

            //determine if the property identifier is in a stream or sub storage
            string propTag = null;
            var propType = NativeMethods.PtUnspecified;

            foreach (var propKey in propKeys)
            {
                if (!propKey.StartsWith("__substg1.0_" + propIdentifier)) continue;
                propTag = propKey.Substring(12, 8);
                propType = ushort.Parse(propKey.Substring(16, 4), NumberStyles.HexNumber);
                break;
            }

            //depending on prop type use method to get property value
            var containerName = "__substg1.0_" + propTag;
            switch (propType)
            {
                case NativeMethods.PtUnspecified:
                    return null;

                case NativeMethods.PtString8:
                    return GetStreamAsString(containerName, Encoding.UTF8);

                case NativeMethods.PtUnicode:
                    return GetStreamAsString(containerName, Encoding.Unicode);

                case NativeMethods.PtBinary:
                    return GetStreamBytes(containerName);

                case NativeMethods.PtObject:
                    return
                        NativeMethods.CloneStorage(
                            _storage.OpenStorage(containerName, IntPtr.Zero,
                                NativeMethods.Stgm.Read | NativeMethods.Stgm.ShareExclusive,
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
            //if no property stream return null
            if (!StreamStatistics.ContainsKey(PropertiesStream))
            {
                return null;
            }

            //get the raw bytes for the property stream
            var propBytes = GetStreamBytes(PropertiesStream);

            //iterate over property stream in 16 byte chunks starting from end of header
            for (var i = _propHeaderSize; i < propBytes.Length; i = i + 16)
            {
                //get property type located in the 1st and 2nd bytes as a unsigned short value
                var propType = BitConverter.ToUInt16(propBytes, i);

                //get property identifer located in 3nd and 4th bytes as a hexdecimal string
                var propIdent = new[] { propBytes[i + 3], propBytes[i + 2] };
                var propIdentString = BitConverter.ToString(propIdent).Replace("-", "");

                //if this is not the property being gotten continue to next property
                if (propIdentString != propIdentifier) continue;

                //depending on prop type use method to get property value
                switch (propType)
                {
                    case NativeMethods.PtI2:
                        return BitConverter.ToInt16(propBytes, i + 8);

                    case NativeMethods.PtLong:
                        return BitConverter.ToInt32(propBytes, i + 8);

                    case NativeMethods.PtSystime:
                        var fileTime = BitConverter.ToInt64(propBytes, i + 8);
                        return DateTime.FromFileTime(fileTime);

                    //default:
                    //throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
                }
            }

            //property not found return null
            return null;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a string.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a string. </returns>
        public string GetMapiPropertyString(string propIdentifier)
        {
            return GetMapiProperty(propIdentifier) as string;
        }

        /// <summary>
        /// Gets the value of the MAPI property as a short.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a short. </returns>
        public Int16 GetMapiPropertyInt16(string propIdentifier)
        {
            return (Int16)GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a integer.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a integer. </returns>
        public int GetMapiPropertyInt32(string propIdentifier)
        {
            return (int)GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// Gets the value of the MAPI property as a byte array.
        /// </summary>
        /// <param name="propIdentifier"> The 4 char hexadecimal prop identifier. </param>
        /// <returns> The value of the MAPI property as a byte array. </returns>
        public byte[] GetMapiPropertyBytes(string propIdentifier)
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
            if (_disposed) return;
            //ensure only disposed once
            _disposed = true;

            //call virtual disposing method to let sub classes clean up
            Disposing();

            //release COM storage object and suppress finalizer
            if (_storage != null)
            {
                ReferenceManager.RemoveItem(_storage);
                Marshal.ReleaseComObject(_storage);
            }
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Gives sub classes the chance to free resources during object disposal.
        /// </summary>
        protected virtual void Disposing()
        {
        }
        #endregion
    }

    // ReSharper restore LocalizableElement
    // ReSharper restore DoNotCallOverridableMethodsInConstructor
}