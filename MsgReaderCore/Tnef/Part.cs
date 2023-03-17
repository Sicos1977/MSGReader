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

using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
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
    private static void ExtractAttachments(TnefReader reader)
    {
        var attachMethod = AttachMethod.ByValue;
        var filter = new BestEncodingFilter();
        var prop = reader.TnefPropertyReader;
        MimePart attachment = null;
        AttachFlags flags;
        var dispose = false;
        string[] mimeType;
        byte[] attachData;
        string text;

        try
        {
            do
            {
                if (reader.AttributeLevel != AttributeLevel.Attachment)
                    break;

                switch (reader.AttributeTag)
                {
                    case AttributeTag.AttachRenderData:
                        attachMethod = AttachMethod.ByValue;
                        if (dispose)
                            attachment.Dispose();
                        attachment = new MimePart();
                        dispose = true;
                        break;
                    case AttributeTag.Attachment:
                        if (attachment is null)
                            break;

                        attachData = null;

                        while (prop.ReadNextProperty())
                            switch (prop.PropertyTag.Id)
                            {
                                case PropertyId.AttachLongFilename:
                                    attachment.FileName = prop.ReadValueAsString();
                                    break;
                                case PropertyId.AttachFilename:
                                    if (attachment.FileName is null)
                                        attachment.FileName = prop.ReadValueAsString();
                                    break;
                                case PropertyId.AttachContentLocation:
                                    attachment.ContentLocation = prop.ReadValueAsUri();
                                    break;
                                case PropertyId.AttachContentBase:
                                    attachment.ContentBase = prop.ReadValueAsUri();
                                    break;
                                case PropertyId.AttachContentId:
                                    text = prop.ReadValueAsString();

                                    var buffer = CharsetUtils.UTF8.GetBytes(text);
                                    var index = 0;

                                    if (ParseUtils.TryParseMsgId(buffer, ref index, buffer.Length, false, false,
                                            out string msgid))
                                        attachment.ContentId = msgid;
                                    break;
                                case PropertyId.AttachDisposition:
                                    text = prop.ReadValueAsString();
                                    if (ContentDisposition.TryParse(text, out ContentDisposition disposition))
                                        attachment.ContentDisposition = disposition;
                                    break;
                                case PropertyId.AttachData:
                                    attachData = prop.ReadValueAsBytes();
                                    break;
                                case PropertyId.AttachMethod:
                                    attachMethod = (AttachMethod)prop.ReadValueAsInt32();
                                    break;
                                case PropertyId.AttachMimeTag:
                                    mimeType = prop.ReadValueAsString().Split('/');
                                    if (mimeType.Length == 2)
                                    {
                                        attachment.ContentType.MediaType = mimeType[0].Trim();
                                        attachment.ContentType.MediaSubtype = mimeType[1].Trim();
                                    }

                                    break;
                                case PropertyId.AttachFlags:
                                    flags = (AttachFlags)prop.ReadValueAsInt32();
                                    if ((flags & AttachFlags.RenderedInBody) != 0)
                                    {
                                        if (attachment.ContentDisposition is null)
                                            attachment.ContentDisposition =
                                                new ContentDisposition(ContentDisposition.Inline);
                                        else
                                            attachment.ContentDisposition.Disposition = ContentDisposition.Inline;
                                    }

                                    break;
                                case PropertyId.AttachSize:
                                    attachment.ContentDisposition ??= new ContentDisposition();

                                    attachment.ContentDisposition.Size = prop.ReadValueAsInt64();
                                    break;
                                case PropertyId.DisplayName:
                                    attachment.ContentType.Name = prop.ReadValueAsString();
                                    break;
                            }

                        if (attachData != null)
                        {
                            var count = attachData.Length;
                            var index = 0;

                            if (attachMethod == AttachMethod.EmbeddedMessage)
                            {
                                attachment.ContentTransferEncoding = ContentEncoding.Base64;
                                attachment = PromoteToTnefPart(attachment);
                                count -= 16;
                                index = 16;
                            }
                            else if (attachment.ContentType.IsMimeType("text", "*"))
                            {
                                filter.Flush(attachData, index, count, out _, out _);
                                attachment.ContentTransferEncoding =
                                    filter.GetBestEncoding(EncodingConstraint.SevenBit);
                                filter.Reset();
                            }
                            else
                            {
                                attachment.ContentTransferEncoding = ContentEncoding.Base64;
                            }

                            attachment.Content = new MimeContent(new MemoryStream(attachData, index, count, false));
                            attachments.Add(attachment);
                            dispose = false;
                        }

                        break;
                    case AttributeTag.AttachCreateDate:
                        if (attachment != null)
                        {
                            if (attachment.ContentDisposition is null)
                                attachment.ContentDisposition = new ContentDisposition();

                            attachment.ContentDisposition.CreationDate = prop.ReadValueAsDateTime();
                        }

                        break;
                    case AttributeTag.AttachModifyDate:
                        if (attachment != null)
                        {
                            if (attachment.ContentDisposition is null)
                                attachment.ContentDisposition = new ContentDisposition();

                            attachment.ContentDisposition.ModificationDate = prop.ReadValueAsDateTime();
                        }

                        break;
                    case AttributeTag.AttachTitle:
                        if (attachment != null && string.IsNullOrEmpty(attachment.FileName))
                            attachment.FileName = prop.ReadValueAsString();
                        break;
                    case AttributeTag.AttachMetaFile:
                        if (attachment is null)
                            break;

                        // TODO: what to do with the meta data?
                        break;
                    case AttributeTag.AttachData:
                        if (attachment is null || attachMethod != AttachMethod.ByValue)
                            break;

                        attachData = prop.ReadValueAsBytes();

                        if (attachment.ContentType.IsMimeType("text", "*"))
                        {
                            filter.Flush(attachData, 0, attachData.Length, out _, out _);
                            attachment.ContentTransferEncoding = filter.GetBestEncoding(EncodingConstraint.SevenBit);
                            filter.Reset();
                        }
                        else
                        {
                            attachment.ContentTransferEncoding = ContentEncoding.Base64;
                        }

                        attachment.Content = new MimeContent(new MemoryStream(attachData, false));
                        attachments.Add(attachment);
                        dispose = false;
                        break;
                }
            } while (reader.ReadNextAttribute());
        }
        finally
        {
            if (dispose)
                attachment.Dispose();
        }
    }

    /// <summary>
    ///     Extract the embedded attachments from the TNEF data.
    /// </summary>
    /// <remarks>
    ///     Parses the TNEF data and extracts all of the embedded file attachments.
    /// </remarks>
    /// <returns>The attachments.</returns>
    /// <exception cref="System.InvalidOperationException">
    ///     The <see cref="MimePart.Content" /> property is <c>null</c>.
    /// </exception>
    /// <exception cref="System.ObjectDisposedException">
    ///     The <see cref="Part" /> has been disposed.
    /// </exception>
    public IEnumerable<MimeEntity> ExtractAttachments()
    {
        var message = ConvertToMessage();
        var body = message.Body;

        // we don't need the message container itself
        message.Body = null;
        message.Dispose();

        if (body is Multipart multipart)
        {
            if (multipart.Count > 0 && multipart[0] is MultipartAlternative alternatives)
            {
                foreach (var alternative in alternatives)
                    yield return alternative;

                alternatives.Clear(false);
                alternatives.Dispose();
                multipart.RemoveAt(0);
            }

            foreach (var part in multipart)
                yield return part;

            multipart.Clear();
            multipart.Dispose();
        }
        else if (body != null)
        {
            yield return body;
        }
    }
}