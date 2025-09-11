using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using MsgReader.Helpers;

namespace MsgReader.Mime.Header;

/// <summary>
///     Utility class that divides a message into a body and a header.<br />
///     The header is then parsed to a strongly typed <see cref="MessageHeader" /> object.
/// </summary>
public static class HeaderExtractor
{
    #region GetHeaders
    /// <summary>
    ///     Extract the headers from the given headers string and gives it back
    ///     as a MessageHeader object
    /// </summary>
    /// <param name="headersString">The string with the header information</param>
    public static MessageHeader GetHeaders(string headersString)
    {
        var headersUnparsedCollection = ExtractHeaders(headersString);
        return new MessageHeader(headersUnparsedCollection);
    }
    #endregion

    #region FindHeaderEndPosition
    /// <summary>
    ///     Find the end of the header section in a byte array.<br />
    ///     The headers have ended when a blank line is found
    /// </summary>
    /// <param name="messageContent">The full message stored as a byte array</param>
    /// <returns>The position of the line just after the header end blank line</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="messageContent" /> is <see langword="null" /></exception>
    private static int FindHeaderEndPosition(byte[] messageContent)
    {
        if (messageContent == null)
            throw new ArgumentNullException(nameof(messageContent));

        // Convert the byte array into a stream
        using Stream stream = StreamHelpers.Manager.GetStream("HeaderExtractor,cs", messageContent, 0, messageContent.Length);
        while (true)
        {
            // Read a line from the stream. We know headers are in US-ASCII
            // therefore it is not problem to read them as such
            var line = StreamUtility.ReadLineAsAscii(stream);

            // The end of headers is signaled when a blank line is found
            // or if the line is null - in which case the email is actually an email with
            // only headers but no e-mail body
            if (string.IsNullOrEmpty(line))
                return (int)stream.Position;
        }
    }
    #endregion

    #region ExtractHeadersAndBody
    /// <summary>
    ///     Extract the header part and body part of a message.<br />
    ///     The headers are then parsed to a strongly typed <see cref="MessageHeader" /> object.
    /// </summary>
    /// <param name="fullRawMessage">The full message in bytes where header and body needs to be extracted from</param>
    /// <param name="loadAttachments">Charge the whole attachments content</param>
    /// <param name="headers">The extracted header parts of the message</param>
    /// <param name="body">The body part of the message</param>
    /// <returns><c>true</c> if message has changed, otherwise <c>false</c></returns>
    /// <exception cref="ArgumentNullException">If <paramref name="fullRawMessage" /> is <see langword="null" /></exception>
    public static bool ExtractHeadersAndBody(byte[] fullRawMessage, bool loadAttachments, out MessageHeader headers, out byte[] body)
    {
        Logger.WriteToLog("Extracting header and body");

        if (fullRawMessage == null)
            throw new ArgumentNullException(nameof(fullRawMessage));

        // Find the end location of the headers
        var endOfHeaderLocation = FindHeaderEndPosition(fullRawMessage);

        // The headers are always in ASCII - therefore we can convert the header part into a string
        // using US-ASCII encoding
        //var headersString = Encoding.ASCII.GetString(fullRawMessage, 0, endOfHeaderLocation);

        // MIME headers should always be ASCII encoded, but sometimes they don't so we read them as UTF8.
        // It should not make any difference if we do it this way because UTF-8 super seeds ASCII encoding
        var headersString = Encoding.UTF8.GetString(fullRawMessage, 0, endOfHeaderLocation);

        // Now parse the headers to a NameValueCollection
        var headersUnparsedCollection = ExtractHeaders(headersString);

        // Use the NameValueCollection to parse it into a strongly-typed MessageHeader header
        headers = new MessageHeader(headersUnparsedCollection);
        if (!loadAttachments && headers.ContentDisposition?.DispositionType == "attachment")
        {

            body = [];
            return true;
        } 
        else
        {
            // Since we know where the headers end, we also know where the body is
            // Copy the body part into the body parameter
            body = new byte[fullRawMessage.Length - endOfHeaderLocation];
            Array.Copy(fullRawMessage, endOfHeaderLocation, body, 0, body.Length);
            return false;
        }
    }
    #endregion

    #region ExtractHeaders
    /// <summary>
    ///     Method that takes a full message and extract the headers from it.
    /// </summary>
    /// <param name="messageContent">
    ///     The message to extract headers from. Does not need the body part. Needs the empty headers
    ///     end line.
    /// </param>
    /// <returns>A collection of Name and Value pairs of headers</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="messageContent" /> is <see langword="null" /></exception>
    private static NameValueCollection ExtractHeaders(string messageContent)
    {
        Logger.WriteToLog("Extracting headers");

        if (messageContent == null)
            throw new ArgumentNullException(nameof(messageContent));

        var headers = new NameValueCollection();

        using var messageReader = new StringReader(messageContent);
        // Read until all headers have ended.
        // The headers ends when an empty line is encountered
        // An empty message might actually not have an empty line, in which
        // case the headers end with null value.
        string line;
        string currentHeaderName = null;
        StringBuilder currentHeaderValue = null;

        while ((line = messageReader.ReadLine()) != null)
        {
            // Check if this is an empty line (end of headers)
            if (string.IsNullOrEmpty(line))
                break;

            // Check if this line starts with whitespace (continuation of previous header)
            // OR if it's a malformed encoded word continuation (starts with "=?")
            if (line.Length > 0 && (char.IsWhiteSpace(line[0]) || 
                (line.StartsWith("=?") && currentHeaderName != null &&
                 (currentHeaderName.Equals("SUBJECT", StringComparison.OrdinalIgnoreCase) ||
                  currentHeaderName.Equals("FROM", StringComparison.OrdinalIgnoreCase) ||
                  currentHeaderName.Equals("TO", StringComparison.OrdinalIgnoreCase)))))
            {
                // This is a continuation of the previous header
                if (currentHeaderValue != null)
                {
                    // For properly formatted headers with leading whitespace, just append after trimming
                    // For malformed headers without leading whitespace, no space needed as they're complete encoded words
                    currentHeaderValue.Append(line.TrimStart());
                }
            }
            else
            {
                // This is a new header. First save the previous one if it exists
                if (currentHeaderName != null && currentHeaderValue != null)
                {
                    if (headers.AllKeys.Contains(currentHeaderName))
                    {
                        var value = headers[currentHeaderName];
                        value += "," + currentHeaderValue;
                        headers[currentHeaderName] = value;
                    }
                    else
                        headers.Add(currentHeaderName, currentHeaderValue.ToString());
                }

                // Now parse the new header
                var header = SeparateHeaderNameAndValue(line);
                currentHeaderName = header.Key;
                currentHeaderValue = new StringBuilder(header.Value);
            }
        }

        // Don't forget to add the last header
        if (currentHeaderName != null && currentHeaderValue != null)
        {
            if (headers.AllKeys.Contains(currentHeaderName))
            {
                var value = headers[currentHeaderName];
                value += "," + currentHeaderValue;
                headers[currentHeaderName] = value;
            }
            else
                headers.Add(currentHeaderName, currentHeaderValue.ToString());
        }

        return headers;
    }
    #endregion

    #region SeparateHeaderNameAndValue
    /// <summary>
    ///     Separate a full header line into a header name and a header value.
    /// </summary>
    /// <param name="rawHeader">The raw header line to be separated</param>
    /// <exception cref="ArgumentNullException">If <paramref name="rawHeader" /> is <see langword="null" /></exception>
    private static KeyValuePair<string, string> SeparateHeaderNameAndValue(string rawHeader)
    {
        if (rawHeader == null)
            throw new ArgumentNullException(nameof(rawHeader));

        var key = string.Empty;
        var value = string.Empty;

        var indexOfColon = rawHeader.IndexOf(':');

        // Check if it is allowed to make substring calls
        if (indexOfColon < 0 || rawHeader.Length < indexOfColon + 1)
            return new KeyValuePair<string, string>(key, value);
        key = rawHeader.Substring(0, indexOfColon).Trim();
        value = rawHeader.Substring(indexOfColon + 1).Trim();

        return new KeyValuePair<string, string>(key, value);
    }
    #endregion
}