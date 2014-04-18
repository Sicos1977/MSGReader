using System;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {        
        /// <summary>
        /// Class represents an attachment
        /// </summary>
        public sealed class Attachment : Storage
        {
            #region Fields
            /// <summary>
            /// Flag to keep track if we already did get the attachment info
            /// </summary>
            private bool _attachmentInfoSet;

            /// <summary>
            /// Contains the attachment data
            /// </summary>
            private byte[] _data;

            /// <summary>
            /// Contains the attachment filename
            /// </summary>
            private string _fileName;

            /// <summary>
            /// Contains the content id
            /// </summary>
            private string _contentId;
            #endregion

            #region Properties
            /// <summary>
            /// Returns the filename of the attachment
            /// </summary>
            public string FileName
            {
                get
                {
                    if (_attachmentInfoSet) return _fileName;
                    GetAttachmentInfo();
                    _attachmentInfoSet = true;
                    return _fileName;
                }
            }

            /// <summary>
            /// Retuns the data
            /// </summary>
            public byte[] Data
            {
                get
                {
                    if (_attachmentInfoSet) return _data;
                    GetAttachmentInfo();
                    _attachmentInfoSet = true;
                    return _data;
                }
            }
            
            /// <summary>
            /// Returns the content id or null when not available
            /// </summary>
            public string ContentId
            {
                get
                {
                    if (_attachmentInfoSet) return _contentId;
                    GetAttachmentInfo();
                    return _contentId;
                }
            }

            /// <summary>
            /// Returns the rendering position or -1 when unknown
            /// </summary>
            public int RenderingPosition
            {
                get
                {
                    var value = GetMapiPropertyInt32(MapiTags.PR_RENDERING_POSITION);
                    if (value == null)
                        return -1;
                    return (int) value;
                }
            }
            #endregion

            #region GetAttachmentInfo
            /// <summary>
            /// Reads the neccesary attachment info as soon as the <see cref="FileName"/> or
            /// <see cref="Data"/> property gets accessed
            /// </summary>
            private void GetAttachmentInfo()
            {
                var fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_FILENAME);

                if (string.IsNullOrEmpty(fileName))
                    fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_FILENAME);

                if (string.IsNullOrEmpty(fileName))
                    fileName = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);

                _fileName = FileManager.RemoveInvalidFileNameChars(fileName);

                _contentId = GetMapiPropertyString(MapiTags.PR_ATTACH_CONTENTID);

                var attachmentMethod = GetMapiPropertyInt32(MapiTags.PR_ATTACH_METHOD);

                switch (attachmentMethod)
                {
                    case MapiTags.ATTACH_OLE:
                        var storage = GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as NativeMethods.IStorage;
                        var attachmentOle = new Attachment(new Storage(storage));
                        try
                        {
                            _data = attachmentOle.GetStreamBytes("CONTENTS");
                            var fileTypeInfo = FileTypeSelector.GetFileTypeFileInfo(_data);
                            _fileName += "." + fileTypeInfo.Extension;
                        }
                        catch (Exception)
                        {
                            _data = null;
                        }
                        break;

                    default:
                        _data = GetMapiPropertyBytes(MapiTags.PR_ATTACH_DATA_BIN);
                        break;
                }

                _attachmentInfoSet = true;
            }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Attachment" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Attachment(Storage message) : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;
            }
            #endregion
        }
    }
}
