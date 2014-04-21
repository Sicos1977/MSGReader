using System;
using System.IO;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    public partial class Storage
    {        
        /// <summary>
        /// Class represents an attachment
        /// </summary>
        public sealed class Attachment : Storage
        {
            #region Fields
            private byte[] _data;
            #endregion

            #region Properties
            /// <summary>
            /// Returns the filename of the attachment
            /// </summary>
            public string FileName { get; private set; }

            /// <summary>
            /// Retuns the data
            /// </summary>
            public byte[] Data
            {
                get { return _data ?? GetMapiPropertyBytes(MapiTags.PR_ATTACH_DATA_BIN); }
            }

            /// <summary>
            /// Returns the content id or null when not available
            /// </summary>
            public string ContentId { get; private set; }

            /// <summary>
            /// Returns the rendering position or -1 when unknown
            /// </summary>
            public int RenderingPosition { get; private set; }

            /// <summary>
            /// True when the attachment is inline
            /// </summary>
            public bool IsInline { get; private set; }

            /// <summary>
            /// True when the attachment is a contact photo. This can only be true
            /// when the <see cref="Storage.Message"/> object is an 
            /// <see cref="Storage.Message.Message.MessageType.Contact"/> object.
            /// </summary>
            public bool IsContactPhoto { get; set; }
            #endregion

            #region ResolveAttachment
            private void ResolveAttachment()
            {
                //The PR_ATTACH_PATHNAME or PR_ATTACH_LONG_PATHNAME property contains a fully qualified path identifying the attachment
                var attachPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_PATHNAME);
                var attachLongPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_PATHNAME);

                // Because we are not sure we can access the files we put everything in a try catch
                try
                {
                    if (attachLongPathName != null)
                    {
                        _data = File.ReadAllBytes(attachLongPathName);
                        return;
                    }

                    if (attachPathName == null) return;
                    _data = File.ReadAllBytes(attachPathName);
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch {}
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

                ContentId = GetMapiPropertyString(MapiTags.PR_ATTACH_CONTENTID);
                IsInline = ContentId != null;

                var isContactPhoto = GetMapiPropertyBool(MapiTags.PR_ATTACHMENT_CONTACTPHOTO);
                if (isContactPhoto == null)
                    IsContactPhoto = false;
                else
                    IsContactPhoto = (bool) isContactPhoto;

                var renderingPosition = GetMapiPropertyInt32(MapiTags.PR_RENDERING_POSITION);
                if (renderingPosition == null)
                    RenderingPosition = -1;
                else
                    RenderingPosition = (int) renderingPosition;

                var fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_FILENAME);

                if (string.IsNullOrEmpty(fileName))
                    fileName = GetMapiPropertyString(MapiTags.PR_ATTACH_FILENAME);

                if (string.IsNullOrEmpty(fileName))
                    fileName = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);

                if (fileName != null)
                    FileName = FileManager.RemoveInvalidFileNameChars(fileName); 
                
                var attachmentMethod = GetMapiPropertyInt32(MapiTags.PR_ATTACH_METHOD);

                switch (attachmentMethod)
                {

                    case MapiTags.ATTACH_BY_REFERENCE:
                    case MapiTags.ATTACH_BY_REF_RESOLVE:
                    case MapiTags.ATTACH_BY_REF_ONLY:
                        ResolveAttachment();
                        break;
                    
                    case MapiTags.ATTACH_OLE:
                        var storage = GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as NativeMethods.IStorage;
                        var attachmentOle = new Attachment(new Storage(storage));
                        _data = attachmentOle.GetStreamBytes("CONTENTS");
                        var fileTypeInfo = FileTypeSelector.GetFileTypeFileInfo(Data);
                        FileName += "." + fileTypeInfo.Extension.ToLower();
                        IsInline = true;
                        break;
                }
            }
            #endregion
        }
    }
}
