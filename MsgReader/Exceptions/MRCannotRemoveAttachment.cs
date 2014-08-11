using System;
using MsgReader.Outlook;

namespace MsgReader.Exceptions
{
    /// <summary>
    /// Raised when it is not possible to remove the <see cref="Storage.Attachment"/> or <see cref="Storage.Message"/> from
    /// the <see cref="Storage.Message"/>
    /// </summary>
    public class MRCannotRemoveAttachment : Exception
    {
        internal MRCannotRemoveAttachment()
        {
        }

        internal MRCannotRemoveAttachment(string message)
            : base(message)
        {
        }

        internal MRCannotRemoveAttachment(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}