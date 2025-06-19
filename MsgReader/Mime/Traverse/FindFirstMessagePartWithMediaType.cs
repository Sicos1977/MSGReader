using System;
using System.Linq;

namespace MsgReader.Mime.Traverse;

///<summary>
/// Finds the first <see cref="MessagePart"/> which have a given MediaType in a depth first traversal.
///</summary>
internal class FindFirstMessagePartWithMediaType : IQuestionAnswerMessageTraverser<string, MessagePart>
{
    #region VisitMessage
    /// <summary>
    /// Finds the first <see cref="MessagePart"/> with the given MediaType
    /// </summary>
    /// <param name="message">The <see cref="Message"/> to start looking in</param>
    /// <param name="question">The MediaType to look for. Case is ignored.</param>
    /// <returns>A <see cref="MessagePart"/> with the given MediaType or <see langword="null"/> if no such <see cref="MessagePart"/> was found</returns>
    public MessagePart VisitMessage(Message message, string question)
    {
        if(message == null)
            throw new ArgumentNullException(nameof(message));

        return VisitMessagePart(message.MessagePart, question);
    }
    #endregion

    #region VisitMessagePart
    /// <summary>
    /// Finds the first <see cref="MessagePart"/> with the given MediaType
    /// </summary>
    /// <param name="messagePart">The <see cref="MessagePart"/> to start looking in</param>
    /// <param name="question">The MediaType to look for. Case is ignored.</param>
    /// <returns>A <see cref="MessagePart"/> with the given MediaType or <see langword="null"/> if no such <see cref="MessagePart"/> was found</returns>
    public MessagePart VisitMessagePart(MessagePart messagePart, string question)
    {
        if(messagePart == null)
            throw new ArgumentNullException(nameof(messagePart));

        if (messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase))
            return messagePart;

        return !messagePart.IsMultiPart
            ? null
            : messagePart.MessageParts.Select(part => VisitMessagePart(part, question))
                .FirstOrDefault(result => result != null);
    }
    #endregion
}
