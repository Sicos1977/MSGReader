using System;

namespace DocumentServices.Modules.Readers.MsgReader.Exceptions
{
    /// <summary>
    /// Raised when the Microsoft Outlook signed message is invalid
    /// </summary>
    public class MRInvalidSignedFile : Exception
    {
        ///
        internal MRInvalidSignedFile()
        {
        }

        internal MRInvalidSignedFile(string message)
            : base(message)
        {
        }

        internal MRInvalidSignedFile(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}