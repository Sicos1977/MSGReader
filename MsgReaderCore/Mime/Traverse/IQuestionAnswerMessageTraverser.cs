namespace MsgReader.Mime.Traverse
{
	/// <summary>
	/// This interface describes a MessageTraverser which is able to traverse a Message structure
	/// and deliver some answer given some question.
	/// </summary>
	/// <typeparam name="TAnswer">This is the type of the answer you want to have delivered.</typeparam>
	/// <typeparam name="TQuestion">This is the type of the question you want to have answered.</typeparam>
	public interface IQuestionAnswerMessageTraverser<in TQuestion, out TAnswer>
	{
		/// <summary>
		/// Call this when you want to apply this traverser on a <see cref="Message"/>.
		/// </summary>
		/// <param name="message">The <see cref="Message"/> which you want to traverse. Must not be <see langword="null"/>.</param>
		/// <param name="question">The question</param>
		/// <returns>An answer</returns>
		TAnswer VisitMessage(Message message, TQuestion question);

		/// <summary>
		/// Call this when you want to apply this traverser on a <see cref="MessagePart"/>.
		/// </summary>
		/// <param name="messagePart">The <see cref="MessagePart"/> which you want to traverse. Must not be <see langword="null"/>.</param>
		/// <param name="question">The question</param>
		/// <returns>An answer</returns>
		TAnswer VisitMessagePart(MessagePart messagePart, TQuestion question);
	}
}