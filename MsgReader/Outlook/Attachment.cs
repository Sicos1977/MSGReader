using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using MsgReader.Exceptions;
using MsgReader.Helpers;
using MsgReader.Localization;

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
    public partial class Storage
    {        
        /// <summary>
        /// Class represents an attachment
        /// </summary>
        public sealed class Attachment : Storage
        {
            #region Fields
            /// <summary>
            /// Containts the data of the attachment as an byte array
            /// </summary>
            private byte[] _data;
            #endregion

            #region Properties
            /// <summary>
            /// The name of the <see cref="Storage.NativeMethods.IStorage"/> stream that containts this attachment
            /// </summary>
            internal string StorageName { get; private set; }

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
            /// <see cref="Storage.Message.MessageType.Contact"/> object.
            /// </summary>
            public bool IsContactPhoto { get; private set; }

            /// <summary>
            /// Returns the date and time when the attachment was created or null
            /// when not available
            /// </summary>
            public DateTime? CreationTime { get; private set; }

            /// <summary>
            /// Returns the date and time when the attachment was last modified or null
            /// when not available
            /// </summary>
            public DateTime? LastModificationTime { get; private set; }

            /// <summary>
            /// Returns <c>true</c> when the attachment is an OLE attachment
            /// </summary>
            public bool OleAttachment { get; private set; }
            #endregion
            
            #region Constructors
            /// <summary>
            /// Creates an attachment object from a <see cref="Mime.MessagePart"/>
            /// </summary>
            /// <param name="attachment"></param>
            internal Attachment(Mime.MessagePart attachment)
            {
                ContentId = attachment.ContentId;
                IsInline = ContentId != null;
                IsContactPhoto = false;
                RenderingPosition = -1;
                _data = attachment.Body;
                FileName = FileManager.RemoveInvalidFileNameChars(attachment.FileName);
                StorageName = null;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Attachment" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            /// <param name="storageName">The name of the <see cref="Storage.NativeMethods.IStorage"/> stream that containts this attachment</param>
            internal Attachment(Storage message, string storageName) : base(message._storage)
            {
                StorageName = storageName;

                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

                CreationTime = GetMapiPropertyDateTime(MapiTags.PR_CREATION_TIME);
                LastModificationTime = GetMapiPropertyDateTime(MapiTags.PR_LAST_MODIFICATION_TIME);

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

                FileName = fileName != null
                    ? FileManager.RemoveInvalidFileNameChars(fileName)
                    : LanguageConsts.NameLessFileName;

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
                        var attachmentOle = new Attachment(new Storage(storage), null);
                        _data = attachmentOle.GetStreamBytes("CONTENTS");
                        if (_data != null)
                        {
                            var fileTypeInfo = FileTypeSelector.GetFileTypeFileInfo(Data);

                            if (string.IsNullOrEmpty(FileName))
                                FileName = fileTypeInfo.Description;

                            FileName += "." + fileTypeInfo.Extension.ToLower();
                        }
                        else
                            // http://www.devsuperpage.com/search/Articles.aspx?G=10&ArtID=142729
                            _data = attachmentOle.GetStreamBytes("\u0002OlePres000");

                        if (_data != null)
                        {
                            try
                            {
                                SaveImageAsPng(40);
                            }
                            catch (ArgumentException)
                            {
                                SaveImageAsPng(0);
                            }
                        }
                        else
                            throw new MRUnknownAttachmentFormat("Can not read the attachment");

                        OleAttachment = true;
                        IsInline = true;
                        break;
                }
            }
            #endregion

            #region ResolveAttachment
            /// <summary>
            /// Tries to resolve an attachment when the <see cref="MapiTags.PR_ATTACH_METHOD"/> is of the type
            /// <see cref="MapiTags.ATTACH_BY_REFERENCE"/>, <see cref="MapiTags.ATTACH_BY_REF_RESOLVE"/> or
            /// <see cref="MapiTags.ATTACH_BY_REF_ONLY"/>
            /// </summary>
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

            #region SaveImageAsPng
            /// <summary>
            /// Tries to save an attachment as a png file with the user specified buffer
            /// </summary>
            private void SaveImageAsPng(int bufferOffset)
            {
                if (bufferOffset > _data.Length)
                    throw new ArgumentOutOfRangeException("bufferOffset", bufferOffset,
                        @"Buffer Offset value cannot be greater than the length of the image byte array!");

                var length = _data.Length - bufferOffset;
                var bytes = new byte[length];
                Buffer.BlockCopy(_data, bufferOffset, bytes, 0, length);
                using (var inputStream = new MemoryStream(bytes))
                using (var image = Image.FromStream(inputStream))
                using (var outputStream = new MemoryStream())
                {
                    image.Save(outputStream, ImageFormat.Png);
                    outputStream.Position = 0;
                    _data = outputStream.ToByteArray();
                    FileName = "ole0.bmp";
                }
            }
            #endregion
        }
    }
}
