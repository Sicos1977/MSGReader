using System;
using System.Collections.Generic;
using MsgReader.Helpers;
using MsgReader.Tnef;
using MsgReader.Tnef.Enums;

namespace MsgReader.Mime.Traverse;

/// <summary>
///     Finds all <see cref="MessagePart" />s which are considered to be attachments
/// </summary>
internal class AttachmentFinder : MultipleMessagePartFinder
{
    #region CaseLeaf
    protected override List<MessagePart> CaseLeaf(MessagePart messagePart)
    {
        if (messagePart == null)
            throw new ArgumentNullException(nameof(messagePart));

        // Maximum space needed is one
        var leafAnswer = new List<MessagePart>(1);

        if (messagePart.IsAttachment)
        {
            if (messagePart.FileName.ToLowerInvariant() == "winmail.dat")
                try
                {
                    var stream = StreamHelpers.Manager.GetStream("AttachmentFinder.CaseLeaf", messagePart.Body, 0,
                        messagePart.Body.Length);
                    using var tnefReader = new TnefReader(stream);
                    {
                        var t = tnefReader.TnefPropertyReader.AttachMethod;
                        while (tnefReader.ReadNextAttribute())
                        {
                            if (tnefReader.AttributeLevel != AttributeLevel.Attachment) continue;

                            var propertyReader = tnefReader.TnefPropertyReader;

                            switch (tnefReader.AttributeTag)
                            {
                                case AttributeTag.Body:
                                    break;
                                case AttributeTag.AttachData:
                                    var data = propertyReader.ReadValueAsBytes();
                                    break;
                                case AttributeTag.AttachTitle:
                                    var attachmentName = propertyReader.ReadValueAsString();
                                    break;
                                case AttributeTag.AttachMetaFile:
                                    break;
                                case AttributeTag.AttachCreateDate:
                                    var attachmentCreateDate = propertyReader.ReadValueAsDateTime();
                                    break;
                                case AttributeTag.AttachModifyDate:
                                    var attachmentModifyDate = propertyReader.ReadValueAsDateTime();
                                    break;
                                case AttributeTag.DateModified:
                                    break;
                                case AttributeTag.AttachTransportFilename:
                                    break;
                                case AttributeTag.Attachment:
                                    break;
                            }
                        }
                    }
                    //tnefReader.
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

            leafAnswer.Add(messagePart);
        }

        return leafAnswer;
    }
    #endregion
}