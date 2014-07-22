using System;

namespace DocumentServices.Modules.Readers.MsgReader.Exceptions
{
    /// <summary>
    /// Raised when the Microsoft Outlook message type or EML is not supported
    /// </summary>
    public class MRFileTypeNotSupported : Exception
    {
        public MRFileTypeNotSupported()
        {
        }

        public MRFileTypeNotSupported(string message)
            : base(message)
        {
        }

        public MRFileTypeNotSupported(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}