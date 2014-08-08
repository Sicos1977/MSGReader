using System;
using System.Collections.Generic;
using System.Linq;

namespace MsgReader.Mime.Traverse
{
	/// <summary>
	/// This is an abstract class which handles traversing of a <see cref="Message"/> tree structure.<br/>
	/// It runs through the message structure using a depth-first traversal.
	/// </summary>
	/// <typeparam name="TAnswer">The answer you want from traversing the message tree structure</typeparam>
	public abstract class AnswerMessageTraverser<TAnswer> : IAnswerMessageTraverser<TAnswer>
    {
        #region VisitMessage
        /// <summary>
	    /// Call this when you want an answer for a full message.
	    /// </summary>
	    /// <param name="message">The message you want to traverse</param>
	    /// <returns>An answer</returns>
	    /// <exception cref="ArgumentNullException">if <paramref name="message"/> is <see langword="null"/></exception>
	    public TAnswer VisitMessage(Message message)
	    {
	        if (message == null)
	            throw new ArgumentNullException("message");

	        return VisitMessagePart(message.MessagePart);
	    }
	    #endregion

        #region VisitMessagePart
        /// <summary>
	    /// Call this method when you want to find an answer for a <see cref="MessagePart"/>
	    /// </summary>
	    /// <param name="messagePart">The <see cref="MessagePart"/> part you want an answer from.</param>
	    /// <returns>An answer</returns>
	    /// <exception cref="ArgumentNullException">if <paramref name="messagePart"/> is <see langword="null"/></exception>
	    public TAnswer VisitMessagePart(MessagePart messagePart)
	    {
	        if (messagePart == null)
	            throw new ArgumentNullException("messagePart");

	        if (!messagePart.IsMultiPart) return CaseLeaf(messagePart);
	        var leafAnswers = new List<TAnswer>(messagePart.MessageParts.Count);
	        leafAnswers.AddRange(messagePart.MessageParts.Select(VisitMessagePart));
	        return MergeLeafAnswers(leafAnswers);
	    }
	    #endregion

		/// <summary>
		/// For a concrete implementation an answer must be returned for a leaf <see cref="MessagePart"/>, which are
		/// MessageParts that are not <see cref="MessagePart.IsMultiPart">MultiParts.</see>
		/// </summary>
		/// <param name="messagePart">The message part which is a leaf and thereby not a MultiPart</param>
		/// <returns>An answer</returns>
		protected abstract TAnswer CaseLeaf(MessagePart messagePart);

		/// <summary>
		/// For a concrete implementation, when a MultiPart <see cref="MessagePart"/> has fetched it's answers from it's children, these
		/// answers needs to be merged. This is the responsibility of this method.
		/// </summary>
		/// <param name="leafAnswers">The answer that the leafs gave</param>
		/// <returns>A merged answer</returns>
		protected abstract TAnswer MergeLeafAnswers(List<TAnswer> leafAnswers);
	}
}