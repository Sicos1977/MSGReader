using System;

namespace DocumentServices.Modules.Readers.MsgReader.Exceptions
{
    /// <summary>
    /// Raised when the Microsoft Outlook message type or EML is not supported
    /// </summary>
    public class MRFileTypeNotSupported : Exception
    {
        internal MRFileTypeNotSupported()
        {
        }

        internal MRFileTypeNotSupported(string message)
            : base(message)
        {
        }

        internal MRFileTypeNotSupported(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}