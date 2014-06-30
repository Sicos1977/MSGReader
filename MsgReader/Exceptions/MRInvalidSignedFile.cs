using System;

namespace DocumentServices.Modules.Readers.MsgReader.Exceptions
{
    /// <summary>
    /// Raised when the Microsoft Outlook signed message is invalid
    /// </summary>
    public class MRInvalidSignedFile : Exception
    {
        public MRInvalidSignedFile()
        {
        }

        public MRInvalidSignedFile(string message)
            : base(message)
        {
        }

        public MRInvalidSignedFile(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}