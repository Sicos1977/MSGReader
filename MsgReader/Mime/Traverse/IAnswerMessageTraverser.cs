namespace MsgReader.Mime.Traverse
{
	/// <summary>
	/// This interface describes a MessageTraverser which is able to traverse a Message hierarchy structure
	/// and deliver some answer.
	/// </summary>
	/// <typeparam name="TAnswer">This is the type of the answer you want to have delivered.</typeparam>
	public interface IAnswerMessageTraverser<out TAnswer>
	{
		/// <summary>
		/// Call this when you want to apply this traverser on a <see cref="Message"/>.
		/// </summary>
		/// <param name="message">The <see cref="Message"/> which you want to traverse. Must not be <see langword="null"/>.</param>
		/// <returns>An answer</returns>
		TAnswer VisitMessage(Message message);

		/// <summary>
		/// Call this when you want to apply this traverser on a <see cref="MessagePart"/>.
		/// </summary>
		/// <param name="messagePart">The <see cref="MessagePart"/> which you want to traverse. Must not be <see langword="null"/>.</param>
		/// <returns>An answer</returns>
		TAnswer VisitMessagePart(MessagePart messagePart);
	}
}