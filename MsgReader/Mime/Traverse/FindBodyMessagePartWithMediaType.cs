using System;

namespace MsgReader.Mime.Traverse
{
	///<summary>
	/// Finds the first <see cref="MessagePart"/> which have a given MediaType in a depth first traversal.
	/// This <see cref="MessagePart"/> will be flagged as the <see cref="MessagePart.IsHtmlBody"/> or
    /// <see cref="MessagePart.IsTextBody"/> according to the searched MediaType ("text/html" or "text/plain")
	///</summary>
    internal class FindBodyMessagePartWithMediaType : IQuestionAnswerMessageTraverser<string, MessagePart>
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
	        if (message == null)
	            throw new ArgumentNullException("message");

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
	        if (messagePart == null)
	            throw new ArgumentNullException("messagePart");

            if (messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase) &&
                (messagePart.ContentDisposition == null ||
                 messagePart.ContentDisposition.DispositionType.ToUpperInvariant() != "ATTACHMENT"))
                return messagePart;
            
            if (!messagePart.IsMultiPart) return null;
	        foreach (var part in messagePart.MessageParts)
	        {
	            var result = VisitMessagePart(part, question);
                if (result != null)
	                return result;
	        }

	        return null;
	    }
	    #endregion
	}
}