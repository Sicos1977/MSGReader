﻿//
// Attachment.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Kees van Spelde. (www.magic-sessions.com)
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
using System.IO;
#if (NETSTANDARD2_0_OR_GREATER)
using Microsoft.Maui.Graphics.Platform;
#else
using System.Drawing.Imaging;
#endif
using MsgReader.Exceptions;
using MsgReader.Helpers;
using MsgReader.Localization;
using MsgReader.Mime;

// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace MsgReader.Outlook;

public partial class Storage
{
    /// <summary>
    ///     Class represents an attachment
    /// </summary>
    public sealed class Attachment : Storage
    {
        #region Fields
        /// <summary>
        ///     contains the data of the attachment as a byte array
        /// </summary>
        private byte[] _data;
        #endregion

        #region Properties
        /// <summary>
        ///     The name of the <see cref="Storage" /> that contains this attachment
        /// </summary>
        internal string StorageName { get; }

        /// <summary>
        ///     Returns the filename of the attachment
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        ///     Returns the data
        /// </summary>
        public byte[] Data => _data ?? GetMapiPropertyBytes(MapiTags.PR_ATTACH_DATA_BIN);

        /// <summary>
        ///     Returns the content id or null when not available
        /// </summary>
        public string ContentId { get; }

        /// <summary>
        ///     Returns the rendering position or -1 when unknown
        /// </summary>
        public int RenderingPosition { get; }

        /// <summary>
        ///     True when the attachment is inline
        /// </summary>
        public bool IsInline { get; internal set; }

        /// <summary>
        ///     True when the attachment is a contact photo. This can only be true
        ///     when the <see cref="Storage.Message" /> object is a <see cref="Storage.Contact" /> object.
        /// </summary>
        public bool IsContactPhoto { get; }

        /// <summary>
        ///     Returns the <see cref="DateTimeOffset"/> when the attachment was created or <c>null</c>
        ///     when not available
        /// </summary>
        /// <remarks>
        ///     Use <see cref="DateTimeOffset.ToLocalTime"/> to get the local time
        /// </remarks>
        public DateTimeOffset? CreationTime { get; }

        /// <summary>
        ///     Returns the <see cref="DateTimeOffset"/>> when the attachment was last modified or <c>null</c>
        ///     when not available
        /// </summary>
        /// <remarks>
        ///     Use <see cref="DateTimeOffset.ToLocalTime"/> to get the local time
        /// </remarks>
        public DateTimeOffset? LastModificationTime { get; }

        /// <summary>
        ///     Returns the MAPI Property Hidden, the value may only exist when it has been set True
        /// </summary>
        public bool Hidden { get; }

        /// <summary>
        ///     Returns <c>true</c> when the attachment is an OLE attachment
        /// </summary>
        public bool OleAttachment { get; }

        /// <summary>
        ///     Returns the mime/type that has been set for the attachment, null will be returned when not set
        /// </summary>
        public string MimeType { get; }
        #endregion

        #region Constructors
        /// <summary>
        ///     Creates an attachment object from a <see cref="Tnef.Attachment" />
        /// </summary>
        /// <param name="attachment"><see cref="Tnef.Attachment"/></param>
        internal Attachment(Tnef.Attachment attachment)
        {
            ContentId = attachment.ContentId;
            IsInline = ContentId != null;
            IsContactPhoto = false;
            RenderingPosition = -1;
            _data = attachment.Body;
            CreationTime = attachment.CreationDate;
            FileName = attachment.FileName;
            MimeType = MimeTypes.GetMimeType(FileName);
            StorageName = null;
        }

        /// <summary>
        ///     Creates an attachment object from a <see cref="Mime.MessagePart" />
        /// </summary>
        /// <param name="attachment"><see cref="MessagePart"/></param>
        internal Attachment(MessagePart attachment)
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
        ///     Initializes a new instance of the <see cref="Storage.Attachment" /> class.
        /// </summary>
        /// <param name="message"> The storage containing the attachment. </param>
        /// <param name="storageName">The name of the <see cref="Storage" /> that contains this attachment</param>
        /// <param name="parentMessage">The parent message (optional, for codepage inheritance)</param>
        internal Attachment(Storage message, string storageName, Storage parentMessage = null) : base(message._rootStorage)
        {
            StorageName = storageName;

            // Inherit the parent message's codepage for proper decoding of PT_STRING8 properties
            if (parentMessage != null)
                _parentMessageCodepage = parentMessage.MessageCodePage;

            GC.SuppressFinalize(message);
            _propHeaderSize = MapiTags.PropertiesStreamHeaderAttachOrRecip;

            CreationTime = GetMapiPropertyDateTimeOffset(MapiTags.PR_CREATION_TIME);
            LastModificationTime = GetMapiPropertyDateTimeOffset(MapiTags.PR_LAST_MODIFICATION_TIME);

            ContentId = GetMapiPropertyString(MapiTags.PR_ATTACH_CONTENTID);
            IsInline = ContentId != null;

            var isHidden = GetMapiPropertyBool(MapiTags.PR_ATTACHMENT_HIDDEN);
            if (isHidden != null)
                Hidden = isHidden.Value;

            var isContactPhoto = GetMapiPropertyBool(MapiTags.PR_ATTACHMENT_CONTACTPHOTO);
            if (isContactPhoto == null)
                IsContactPhoto = false;
            else
                IsContactPhoto = (bool)isContactPhoto;

            var renderingPosition = GetMapiPropertyInt32(MapiTags.PR_RENDERING_POSITION);
            if (renderingPosition == null)
                RenderingPosition = -1;
            else
                RenderingPosition = (int)renderingPosition;

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
                    var storage = GetMapiProperty(MapiTags.PR_ATTACH_DATA_BIN) as OpenMcdf.Storage;
                    var attachmentOle = new Attachment(new Storage(storage), null);

                    _data = attachmentOle.GetStreamBytes("CONTENTS");
                    if (_data != null)
                    {
                        var fileTypeInfo = FileTypeSelector.GetFileTypeFileInfo(Data);

                        if (string.IsNullOrEmpty(FileName))
                            FileName = fileTypeInfo.Description;

                        FileName += fileTypeInfo.Extension.ToLower();
                    }
                    else
                    // http://www.devsuperpage.com/search/Articles.aspx?G=10&ArtID=142729
                    {
                        _data = attachmentOle.GetStreamBytes("\u0002OlePres000");
                    }

                    if (_data != null)
                    {
                        int[] bufferOffsets = [40, 0];
                        foreach (var i in bufferOffsets)
                        {
                            try
                            {
                                SaveImageAsPng(i);
                            }
                            catch (Exception)
                            {
                               //Not an image or not convertible; but hope to still read the email contents. 
                            }
                        }
                    }
                    else
                        throw new MRUnknownAttachmentFormat("Can not read the attachment");

                    OleAttachment = true;
                    IsInline = true;
                    break;
            }

            MimeType = GetMapiPropertyString(MapiTags.PR_ATTACH_MIME_TAG);
        }
        #endregion

        #region ResolveAttachment
        /// <summary>
        ///     Tries to resolve an attachment when the <see cref="MapiTags.PR_ATTACH_METHOD" /> is of the type
        ///     <see cref="MapiTags.ATTACH_BY_REFERENCE" />, <see cref="MapiTags.ATTACH_BY_REF_RESOLVE" /> or
        ///     <see cref="MapiTags.ATTACH_BY_REF_ONLY" />
        /// </summary>
        private void ResolveAttachment()
        {
            //The PR_ATTACH_PATHNAME or PR_ATTACH_LONG_PATHNAME property contains a fully qualified path identifying the attachment
            var attachPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_PATHNAME);
            var attachLongPathName = GetMapiPropertyString(MapiTags.PR_ATTACH_LONG_PATHNAME);

            // Because we are not sure that we can access the files we put everything in a try catch
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
            catch
            {
            }
        }
        #endregion

        #region SaveImageAsPng
        /// <summary>
        ///     Tries to save an attachment as a png file with the user specified buffer
        /// </summary>
        private void SaveImageAsPng(int bufferOffset)
        {
            if (bufferOffset > _data.Length)
                throw new ArgumentOutOfRangeException(nameof(bufferOffset), bufferOffset, @"Buffer offset value cannot be greater than the length of the image byte array!");

            var length = _data.Length - bufferOffset;
            var bytes = new byte[length];
            Buffer.BlockCopy(_data, bufferOffset, bytes, 0, length);

            using var inputStream = StreamHelpers.Manager.GetStream("Attachment.cs", bytes, 0, bytes.Length);
            using var outputStream = StreamHelpers.Manager.GetStream();

#if (NETSTANDARD2_0_OR_GREATER)
            
            using var image = PlatformImage.FromStream(inputStream);
            image.Save(outputStream);
            outputStream.Position = 0;
            _data = outputStream.ToArray();

#else

            using var image = System.Drawing.Image.FromStream(inputStream);
            image.Save(outputStream, ImageFormat.Png);
            outputStream.Position = 0;
            _data = outputStream.ToArray();

#endif

            FileName = $"ole{(RenderingPosition != -1 ? RenderingPosition : 0)}.png";
        }
        #endregion
    }
}