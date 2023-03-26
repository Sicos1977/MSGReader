using System;
using System.Collections.Generic;
using MsgReader.Helpers;
using MsgReader.Tnef;

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
            {
                try
                {
                    Logger.WriteToLog("Found winmail.dat attachment, trying to get attachments from it");
                    var stream = StreamHelpers.Manager.GetStream("AttachmentFinder.CaseLeaf", messagePart.Body, 0, messagePart.Body.Length);
                    using var tnefReader = new TnefReader(stream);
                    {
                        var attachments = Part.ExtractAttachments(tnefReader);
                        var count = attachments.Count;
                        if (count > 0)
                        {
                            Logger.WriteToLog($"Found {count} attachment{(count == 1 ? string.Empty : "s")}, removing winmail.dat and adding {(count == 1 ? "this attachment" : "these attachments")}");

                            foreach (var attachment in attachments)
                            {
                                var temp = new MessagePart(attachment);
                                leafAnswer.Add(temp);
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    Logger.WriteToLog($"Could not parse winmail.dat attachment, error: {ExceptionHelpers.GetInnerException(exception)}");
                    leafAnswer.Add(messagePart);
                }
            }
            else
                leafAnswer.Add(messagePart);
        }

        return leafAnswer;
    }
    #endregion
}