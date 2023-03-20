//
// TnefPart.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
//
// Refactoring to the code done by Kees van Spelde so that it works in this project
// Copyright (c) 2023 Kees van Spelde <sicos2002@hotmail.com>
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
using System.Net.Mime;
using MsgReader.Helpers;
using MsgReader.Mime.Header;
using MsgReader.Tnef.Enums;

namespace MsgReader.Tnef;

/// <summary>
///     A MIME part containing Microsoft TNEF data.
/// </summary>
/// <remarks>
///     <para>Represents an application/ms-tnef or application/vnd.ms-tnef part.</para>
///     <para>
///         TNEF (Transport Neutral Encapsulation Format) attachments are most often
///         sent by Microsoft Outlook clients.
///     </para>
/// </remarks>
internal class Part
{
    #region ExtractHtmlBody
    internal static Attachment ExtractHtmlBody(TnefReader reader)
    {
        var attachment = new Attachment();
        var property = reader.TnefPropertyReader;

        while (property.ReadNextProperty())
        {
            switch (property.PropertyTag.Id)
            {
                case PropertyId.CreationTime:
                    attachment.CreationDate = property.ReadValueAsDateTime();
                    break;

                case PropertyId.LastModificationTime:
                    attachment.ModificationDate = property.ReadValueAsDateTime();
                    break;

                case PropertyId.Subject:
                    attachment.FileName = FileManager.RemoveInvalidFileNameChars(property.ReadValueAsString()) + ".html";
                    break;

                case PropertyId.BodyHtml:
                    if (property.PropertyTag.ValueTnefType is PropertyType.String8 or PropertyType.Unicode or PropertyType.Binary)
                    {
                        attachment.Encoding = property.GetMessageEncoding();
                        attachment.Type = AttachmentType.Html;
                        attachment.ContentType = new ContentType("text/html");
                        attachment.Body = property.ReadValueAsBytes();
                    }

                    break;

                //case PropertyId.RtfCompressed:
                //    break;

                //case PropertyId.Body:
                //    break;
            }
        }

        return attachment;
    }
    #endregion

    #region ExtractAttachments
    /// <summary>
    ///     Extract the attachment from a TNEF winmail.dat file
    /// </summary>
    /// <param name="reader"><see cref="TnefReader"/></param>
    /// <returns></returns>
    internal static List<Attachment> ExtractAttachments(TnefReader reader)
    {
        var attachments = new List<Attachment>();
        Attachment attachment = null;
        var attachMethod = AttachMethod.ByValue;
        var property = reader.TnefPropertyReader;

        do
        {
            switch (reader.AttributeLevel)
            {
                case AttributeLevel.Message:
                    var htmlAttachment = ExtractHtmlBody(reader);
                    if (htmlAttachment.Body != null) 
                        attachments.Add(htmlAttachment);
                    break;

                case AttributeLevel.Attachment:
                {
                    byte[] attachmentData;

                    switch (reader.AttributeTag)
                    {
                        case AttributeTag.AttachRenderData:
                            attachMethod = AttachMethod.ByValue;
                            attachment = new Attachment { ContentType = new ContentType() };
                            break;

                        case AttributeTag.Attachment:
                            if (attachment is null)
                                break;

                            attachmentData = null;

                            while (property.ReadNextProperty())
                            {
                                string text;
                                switch (property.PropertyTag.Id)
                                {
                                    case PropertyId.AttachLongFilename:
                                        attachment.FileName = property.ReadValueAsString();
                                        break;

                                    case PropertyId.AttachFilename:
                                        attachment.FileName ??= property.ReadValueAsString();
                                        break;

                                    case PropertyId.AttachContentLocation:
                                        attachment.ContentLocation = property.ReadValueAsUri();
                                        break;

                                    case PropertyId.AttachContentBase:
                                        attachment.ContentBase = property.ReadValueAsUri();
                                        break;

                                    case PropertyId.AttachContentId:
                                        text = property.ReadValueAsString();
                                        attachment.ContentId = text;
                                        break;

                                    case PropertyId.AttachDisposition:
                                        text = property.ReadValueAsString();
                                        attachment.ContentDisposition = new ContentDisposition(text);
                                        break;

                                    case PropertyId.AttachData:
                                        attachmentData = property.ReadValueAsBytes();
                                        break;

                                    case PropertyId.AttachMethod:
                                        attachMethod = (AttachMethod)property.ReadValueAsInt32();
                                        break;

                                    case PropertyId.AttachMimeTag:
                                        var mimeType = property.ReadValueAsString();
                                        attachment.ContentType = new ContentType(mimeType);

                                        break;

                                    case PropertyId.AttachFlags:

                                        var flags = (AttachFlags)property.ReadValueAsInt32();

                                        if ((flags & AttachFlags.RenderedInBody) != 0)
                                        {
                                            attachment.ContentDisposition ??= new ContentDisposition();
                                            attachment.ContentDisposition.Inline = true;
                                        }

                                        break;

                                    case PropertyId.AttachSize:
                                        attachment.ContentDisposition ??= new ContentDisposition();
                                        attachment.ContentDisposition.Size = property.ReadValueAsInt64();
                                        break;

                                    case PropertyId.DisplayName:
                                        attachment.ContentType.Name = property.ReadValueAsString();
                                        break;
                                }
                            }

                            if (attachmentData != null)
                            {
                                var count = attachmentData.Length;
                                var index = 0;

                                //if (attachMethod == AttachMethod.EmbeddedMessage)
                                //{
                                //    attachment.ContentTransferEncoding = ContentTransferEncoding.Base64;
                                //    //attachment = PromoteToTnefPart(attachment);
                                //    count -= 16;
                                //    index = 16;
                                //}
                                //else if (attachment.ContentType.MediaType.StartsWith("text/"))
                                //{
                                //    //filter.Flush(attachData, index, count, out _, out _);
                                //    //attachment.ContentTransferEncoding =
                                //    //    filter.GetBestEncoding(EncodingConstraint.SevenBit);
                                //    //filter.Reset();
                                //}
                                //else
                                attachment.ContentTransferEncoding = ContentTransferEncoding.Base64;

                                attachment.Body = new byte[count];
                                Array.Copy(attachmentData, index, attachment.Body, 0, count);
                                attachments.Add(attachment);
                            }

                            break;

                        case AttributeTag.AttachCreateDate:
                            if (attachment != null)
                            {
                                attachment.ContentDisposition ??= new ContentDisposition();
                                attachment.ContentDisposition.CreationDate = property.ReadValueAsDateTime();
                            }

                            break;

                        case AttributeTag.AttachModifyDate:
                            if (attachment != null)
                            {
                                attachment.ContentDisposition ??= new ContentDisposition();
                                attachment.ContentDisposition.ModificationDate = property.ReadValueAsDateTime();
                            }

                            break;
                        case AttributeTag.AttachTitle:
                            if (attachment != null && string.IsNullOrEmpty(attachment.FileName))
                                attachment.FileName = property.ReadValueAsString();
                            break;

                        case AttributeTag.AttachMetaFile:
                            break;

                        case AttributeTag.AttachData:
                            if (attachment is null || attachMethod != AttachMethod.ByValue)
                                break;

                            attachmentData = property.ReadValueAsBytes();

                            //if (attachment.ContentType.MediaType.StartsWith("text/"))
                            //{
                            //    //filter.Flush(attachData, 0, attachData.Length, out _, out _);
                            //    //attachment.ContentTransferEncoding = filter.GetBestEncoding(EncodingConstraint.SevenBit);
                            //    //filter.Reset();
                            //}
                            //else
                            attachment.ContentTransferEncoding = ContentTransferEncoding.Base64;
                            attachment.Body = attachmentData;
                            attachments.Add(attachment);
                            break;
                    }

                }
                break;
            }
        } while (reader.ReadNextAttribute());

        return attachments;
    }
    #endregion
}