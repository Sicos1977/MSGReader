using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;

namespace DocumentServices.Modules.Readers.MsgReader.Header
{
	///<summary>
	/// Utility class that divides a message into a body and a header.<br/>
	/// The header is then parsed to a strongly typed <see cref="MessageHeader"/> object.
	///</summary>
	internal static class HeaderExtractor
    {
        #region GetHeaders
        /// <summary>
        /// Extract the headers from the given headers string and gives it back
        /// as a MessageHeader object
        /// </summary>
        /// <param name="headersString">The string with the header information</param>
        public static MessageHeader GetHeaders(string headersString)
        {
            var headersUnparsedCollection = ExtractHeaders(headersString);
            return new MessageHeader(headersUnparsedCollection);              
        }
        #endregion

        #region ExtractHeaders
        /// <summary>
		/// Method that takes a full message and extract the headers from it.
		/// </summary>
		/// <param name="messageContent">The message to extract headers from. Does not need the body part. Needs the empty headers end line.</param>
		/// <returns>A collection of Name and Value pairs of headers</returns>
		/// <exception cref="ArgumentNullException">If <paramref name="messageContent"/> is <see langword="null"/></exception>
		private static NameValueCollection ExtractHeaders(string messageContent)
		{
			if(messageContent == null)
				throw new ArgumentNullException("messageContent");

			var headers = new NameValueCollection();

			using (var messageReader = new StringReader(messageContent))
			{
				// Read until all headers have ended.
				// The headers ends when an empty line is encountered
				// An empty message might actually not have an empty line, in which
				// case the headers end with null value.
				string line;
				while (!string.IsNullOrEmpty(line = messageReader.ReadLine()))
				{
					// Split into name and value
					var header = SeparateHeaderNameAndValue(line);

					// First index is header name
					var headerName = header.Key;

					// Second index is the header value.
					// Use a StringBuilder since the header value may be continued on the next line
					var headerValue = new StringBuilder(header.Value);

					// Keep reading until we would hit next header
					// This if for handling multi line headers
					while (IsMoreLinesInHeaderValue(messageReader))
					{
						// Unfolding is accomplished by simply removing any CRLF
						// that is immediately followed by WSP
						// This was done using ReadLine (it discards CRLF)
						// See http://tools.ietf.org/html/rfc822#section-3.1.1 for more information
						var moreHeaderValue = messageReader.ReadLine();

						// If this exception is ever raised, there is an serious algorithm failure
						// IsMoreLinesInHeaderValue does not return true if the next line does not exist
						// This check is only included to stop the nagging "possibly null" code analysis hint
						if (moreHeaderValue == null)
							throw new ArgumentException("This will never happen");

						// Simply append the line just read to the header value
						headerValue.Append(moreHeaderValue);
					}

					// Now we have the name and full value. Add it
					headers.Add(headerName, headerValue.ToString());
				}
			}

			return headers;
		}
        #endregion

        #region IsMoreLinesInHeaderValue
        /// <summary>
		/// Check if the next line is part of the current header value we are parsing by
		/// peeking on the next character of the <see cref="TextReader"/>.<br/>
		/// This should only be called while parsing headers.
		/// </summary>
		/// <param name="reader">The reader from which the header is read from</param>
		/// <returns><see langword="true"/> if multi-line header. <see langword="false"/> otherwise</returns>
		private static bool IsMoreLinesInHeaderValue(TextReader reader)
		{
			var peek = reader.Peek();
			if (peek == -1)
				return false;

			var peekChar = (char)peek;

			// A multi line header must have a whitespace character
			// on the next line if it is to be continued
			return peekChar == ' ' || peekChar == '\t';
		}
        #endregion

        #region SeparateHeaderNameAndValue
        /// <summary>
		/// Separate a full header line into a header name and a header value.
		/// </summary>
		/// <param name="rawHeader">The raw header line to be separated</param>
		/// <exception cref="ArgumentNullException">If <paramref name="rawHeader"/> is <see langword="null"/></exception>
		internal static KeyValuePair<string, string> SeparateHeaderNameAndValue(string rawHeader)
		{
			if (rawHeader == null)
				throw new ArgumentNullException("rawHeader");

			var key = string.Empty;
			var value = string.Empty;

			var indexOfColon = rawHeader.IndexOf(':');

			// Check if it is allowed to make substring calls
			if (indexOfColon >= 0 && rawHeader.Length >= indexOfColon + 1)
			{
				key = rawHeader.Substring(0, indexOfColon).Trim();
				value = rawHeader.Substring(indexOfColon + 1).Trim();
			}

			return new KeyValuePair<string, string>(key, value);
        }
        #endregion
    }
}