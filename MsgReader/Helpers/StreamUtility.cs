using System;
using System.IO;
using System.Text;

namespace MsgReader.Helpers
{
	/// <summary>
	/// Utility to help reading bytes and strings of a <see cref="Stream"/>
	/// </summary>
	internal static class StreamUtility
    {
        #region ReadLineAsBytes
        /// <summary>
	    /// Read a line from the stream.
	    /// A line is interpreted as all the bytes read until a CRLF or LF is encountered.<br/>
	    /// CRLF pair or LF is not included in the string.
	    /// </summary>
	    /// <param name="stream">The stream from which the line is to be read</param>
	    /// <returns>A line read from the stream returned as a byte array or <see langword="null"/> if no bytes were readable from the stream</returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="stream"/> is <see langword="null"/></exception>
	    public static byte[] ReadLineAsBytes(Stream stream)
	    {
	        if (stream == null)
	            throw new ArgumentNullException("stream");

	        using (var memoryStream = new MemoryStream())
	        {
	            while (true)
	            {
	                var justRead = stream.ReadByte();
	                if (justRead == -1 && memoryStream.Length > 0)
	                    break;

	                // Check if we started at the end of the stream we read from
	                // and we have not read anything from it yet
	                if (justRead == -1 && memoryStream.Length == 0)
	                    return null;

	                var readChar = (char) justRead;

	                // Do not write \r or \n
	                if (readChar != '\r' && readChar != '\n')
	                    memoryStream.WriteByte((byte) justRead);

	                // Last point in CRLF pair
	                if (readChar == '\n')
	                    break;
	            }

	            return memoryStream.ToArray();
	        }
	    }
	    #endregion
        
        #region ReadLineAsAscii
        /// <summary>
	    /// Read a line from the stream. <see cref="ReadLineAsBytes"/> for more documentation.
	    /// </summary>
	    /// <param name="stream">The stream to read from</param>
	    /// <returns>A line read from the stream or <see langword="null"/> if nothing could be read from the stream</returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="stream"/> is <see langword="null"/></exception>
	    public static string ReadLineAsAscii(Stream stream)
	    {
	        var readFromStream = ReadLineAsBytes(stream);
	        return readFromStream != null ? Encoding.ASCII.GetString(readFromStream) : null;
	    }
	    #endregion
	}
}