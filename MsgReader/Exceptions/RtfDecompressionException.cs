using System;

namespace MsgReader.Exceptions;

/// <summary>
///     Indicates that a compressed RTF stream is invalid or cannot be decompressed.
/// </summary>
public class RtfDecompressionException : FormatException
{
    /// <summary>
    ///     Creates an exception describing the compressed RTF format error.
    /// </summary>
    /// <param name="message">The error message.</param>
    public RtfDecompressionException(string message) : base(message)
    {
    }

    /// <summary>
    ///     Creates an exception describing the compressed RTF format error and its cause.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The underlying exception.</param>
    public RtfDecompressionException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
