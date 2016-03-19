using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
using System.Text;
using MsgReader.Helpers;
using MsgReader.Localization;
using MsgReader.Mime.Decode;
using MsgReader.Mime.Header;

namespace MsgReader.Mime
{
	/// <summary>
	/// A MessagePart is a part of an email message used to describe the whole email parse tree.<br/>
	/// <br/>
	/// <b>Email messages are tree structures</b>:<br/>
	/// Email messages may contain large tree structures, and the MessagePart are the nodes of the this structure.<br/>
	/// A MessagePart may either be a leaf in the structure or a internal node with links to other MessageParts.<br/>
	/// The root of the message tree is the <see cref="Message"/> class.<br/>
	/// <br/>
	/// <b>Leafs</b>:<br/>
	/// If a MessagePart is a leaf, the part is not a <see cref="IsMultiPart">MultiPart</see> message.<br/>
	/// Leafs are where the contents of an email are placed.<br/>
	/// This includes, but is not limited to: attachments, text or images referenced from HTML.<br/>
	/// The content of an attachment can be fetched by using the <see cref="Body"/> property.<br/>
	/// If you want to have the text version of a MessagePart, use the <see cref="GetBodyAsText"/> method which will<br/>
	/// convert the <see cref="Body"/> into a string using the encoding the message was sent with.<br/>
	/// <br/>
	/// <b>Internal nodes</b>:<br/>
	/// If a MessagePart is an internal node in the email tree structure, then the part is a <see cref="IsMultiPart">MultiPart</see> message.<br/>
	/// The <see cref="MessageParts"/> property will then contain links to the parts it contain.<br/>
	/// The <see cref="Body"/> property of the MessagePart will not be set.<br/>
	/// <br/>
	/// See the example for a parsing example.<br/>
	/// This class cannot be instantiated from outside the library.
	/// </summary>
	/// <example>
	/// This example illustrates how the message parse tree looks like given a specific message<br/>
	/// <br/>
	/// The message source in this example is:<br/>
	/// <code>
	/// MIME-Version: 1.0
	///	Content-Type: multipart/mixed; boundary="frontier"
	///	
	///	This is a message with multiple parts in MIME format.
	///	--frontier
	/// Content-Type: text/plain
	///	
	///	This is the body of the message.
	///	--frontier
	///	Content-Type: application/octet-stream
	///	Content-Transfer-Encoding: base64
	///	
	///	PGh0bWw+CiAgPGHLYWQ+CiAgPC9oZWFkPgogIDxib2R5PgogICAgPHA+VGhpcyBpcyB0aGUg
	///	Ym9keSBvZiB0aGUgbWVzc2FnZS48L3A+CiAgPC9ib2R5Pgo8L2h0bWw+Cg==
	///	--frontier--
	/// </code>
	/// The tree will look as follows, where the content-type media type of the message is listed<br/>
	/// <code>
	/// - Message root
	///   - multipart/mixed MessagePart
	///     - text/plain MessagePart
	///     - application/octet-stream MessagePart
	/// </code>
	/// It is possible to have more complex message trees like the following:<br/>
	/// <code>
	/// - Message root
	///   - multipart/mixed MessagePart
	///     - text/plain MessagePart
	///     - text/plain MessagePart
	///     - multipart/parallel
	///       - audio/basic
	///       - image/tiff
	///     - text/enriched
	///     - message/rfc822
	/// </code>
	/// But it is also possible to have very simple message trees like:<br/>
	/// <code>
	/// - Message root
	///   - text/plain
	/// </code>
	/// </example>
	public class MessagePart
	{
		#region Properties
		/// <summary>
		/// The Content-Type header field.<br/>
		/// <br/>
		/// If not set, the ContentType is created by the default "text/plain; charset=us-ascii" which is
		/// defined in <a href="http://tools.ietf.org/html/rfc2045#section-5.2">RFC 2045 section 5.2</a>.<br/>
		/// <br/>
		/// If set, the default is overridden.
		/// </summary>
		public ContentType ContentType { get; private set; }

		/// <summary>
		/// A human readable description of the body<br/>
		/// <br/>
		/// <see langword="null"/> if no Content-Description header was present in the message.<br/>
		/// </summary>
		public string ContentDescription { get; private set; }

		/// <summary>
		/// This header describes the Content encoding during transfer.<br/>
		/// <br/>
		/// If no Content-Transfer-Encoding header was present in the message, it is set
		/// to the default of <see cref="Header.ContentTransferEncoding.SevenBit">SevenBit</see> in accordance to the RFC.
		/// </summary>
		/// <remarks>See <a href="http://tools.ietf.org/html/rfc2045#section-6">RFC 2045 section 6</a> for details</remarks>
		public ContentTransferEncoding ContentTransferEncoding { get; private set; }

		/// <summary>
		/// ID of the content part (like an attached image). Used with MultiPart messages.<br/>
		/// <br/>
		/// <see langword="null"/> if no Content-ID header field was present in the message.
		/// </summary>
		public string ContentId { get; private set; }

		/// <summary>
		/// Used to describe if a <see cref="MessagePart"/> is to be displayed or to be though of as an attachment.<br/>
		/// Also contains information about filename if such was sent.<br/>
		/// <br/>
		/// <see langword="null"/> if no Content-Disposition header field was present in the message
		/// </summary>
		public ContentDisposition ContentDisposition { get; private set; }

		/// <summary>
		/// This is the encoding used to parse the message body if the <see cref="MessagePart"/><br/>
		/// is not a MultiPart message. It is derived from the <see cref="ContentType"/> character set property.
		/// </summary>
		public Encoding BodyEncoding { get; private set; }

		/// <summary>
		/// This is the parsed body of this <see cref="MessagePart"/>.<br/>
		/// It is parsed in that way, if the body was ContentTransferEncoded, it has been decoded to the
		/// correct bytes.<br/>
		/// <br/>
		/// It will be <see langword="null"/> if this <see cref="MessagePart"/> is a MultiPart message.<br/>
		/// Use <see cref="IsMultiPart"/> to check if this <see cref="MessagePart"/> is a MultiPart message.
		/// </summary>
		public byte[] Body { get; private set; }

        /// <summary>
        /// This will be set to true if this is the first found Text <see cref="MessagePart"/>. This way it
        /// indicates that this is the text variant of the E-mail body.
        /// </summary>
        internal bool IsTextBody { get;  set; }

        /// <summary>
        /// This will be set to true if this is the first found Html <see cref="MessagePart"/>. This way it
        /// indicates that this is the html variant of the E-mail body.
        /// </summary>
        internal bool IsHtmlBody { get; set; }

		/// <summary>
		/// Describes if this <see cref="MessagePart"/> is a MultiPart message<br/>
		/// <br/>
		/// The <see cref="MessagePart"/> is a MultiPart message if the <see cref="ContentType"/> media type property starts with "multipart/"
		/// </summary>
		public bool IsMultiPart
		{
		    get { return ContentType.MediaType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase); }
		}

		/// <summary>
		/// A <see cref="MessagePart"/> is considered to be holding text in it's body if the MediaType
		/// starts either "text/" or is equal to "message/rfc822"
		/// </summary>
		public bool IsText
		{
			get
			{
				var mediaType = ContentType.MediaType;
			    return mediaType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
			           mediaType.Equals("message/rfc822", StringComparison.OrdinalIgnoreCase);
			}
		}

        /// <summary>
        /// A <see cref="MessagePart"/> is considered to be an inline attachment, if<br/>
        /// it is has the <see cref="ContentDisposition"/> Inline set to <c>True</c>
        /// </summary>
        public bool IsInline
	    {
	        get { return ContentDisposition != null && ContentDisposition.Inline; }
	    }

		/// <summary>
		/// A <see cref="MessagePart"/> is considered to be an attachment, if<br/>
		/// - it is not holding <see cref="IsText">text</see> and is not a <see cref="IsMultiPart">MultiPart</see> message<br/>
		/// or<br/>
		/// - it has a Content-Disposition header that says it is an attachment
		/// </summary>
		public bool IsAttachment
		{
			get
			{
				// Inline is the opposite of attachment
			    if (IsHtmlBody)
			        return false;

			    if (IsTextBody)
			        return false;

			    return !IsMultiPart;
			}
		}

		/// <summary>
		/// This is a convenient-property for figuring out a FileName for this <see cref="MessagePart"/>.<br/>
		/// If the <see cref="MessagePart"/> is a MultiPart message, then it makes no sense to try to find a FileName.<br/>
		/// <br/>
		/// The FileName can be specified in the <see cref="ContentDisposition"/>, <see cref="ContentType"/> or 
		/// <see cref="ContentDescription"/> properties.<br/>
		/// If none of these places two places tells about the FileName, a default is returned.
		/// </summary>
		public string FileName { get; private set; }

		/// <summary>
		/// If this <see cref="MessagePart"/> is a MultiPart message, then this property
		/// has a list of each of the Multiple parts that the message consists of.<br/>
		/// <br/>
		/// It is <see langword="null"/> if it is not a MultiPart message.<br/>
		/// Use <see cref="IsMultiPart"/> to check if this <see cref="MessagePart"/> is a MultiPart message.
		/// </summary>
		public List<MessagePart> MessageParts { get; private set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Used to construct the topmost message part
		/// </summary>
		/// <param name="rawBody">The body that needs to be parsed</param>
		/// <param name="headers">The headers that should be used from the message</param>
		/// <exception cref="ArgumentNullException">If <paramref name="rawBody"/> or <paramref name="headers"/> 
		/// is <see langword="null"/></exception>
		internal MessagePart(byte[] rawBody, MessageHeader headers)
		{
			if(rawBody == null)
				throw new ArgumentNullException("rawBody");
			
			if(headers == null)
				throw new ArgumentNullException("headers");

			ContentType = headers.ContentType;
			ContentDescription = headers.ContentDescription;
			ContentTransferEncoding = headers.ContentTransferEncoding;
			ContentId = headers.ContentId;
			ContentDisposition = headers.ContentDisposition;

            FileName = FindFileName(rawBody, headers, LanguageConsts.NameLessFileName);
			BodyEncoding = ParseBodyEncoding(ContentType.CharSet);

			ParseBody(rawBody);
		}
		#endregion

        #region ParseBodyEncoding
        /// <summary>
		/// Parses a character set into an encoding
		/// </summary>
		/// <param name="characterSet">The character set that needs to be parsed. <see langword="null"/> is allowed.</param>
		/// <returns>The encoding specified by the <paramref name="characterSet"/> parameter, or ASCII if the character set 
		/// was <see langword="null"/> or empty</returns>
		private static Encoding ParseBodyEncoding(string characterSet)
		{
			// Default encoding in Mime messages is US-ASCII
			var encoding = Encoding.ASCII;

			// If the character set was specified, find the encoding that the character
			// set describes, and use that one instead
			if (!string.IsNullOrEmpty(characterSet))
				encoding = EncodingFinder.FindEncoding(characterSet);

			return encoding;
		}
        #endregion

        #region FindFileName
        /// <summary>
        /// Figures out the filename of this message part.
        /// <see cref="FileName"/> property.
        /// </summary>
        /// <param name="rawBody">The body that needs to be parsed</param>
        /// <param name="headers">The headers that should be used from the message</param>
        /// <param name="defaultName">The default filename to use, if no other could be found</param>
        /// <returns>The filename found, or the default one if not such filename could be found in the headers</returns>
        /// <exception cref="ArgumentNullException">if <paramref name="headers"/> is <see langword="null"/></exception>
        private static string FindFileName(byte[] rawBody, MessageHeader headers, string defaultName)
        {
            if (headers == null)
                throw new ArgumentNullException("headers");

            if (headers.ContentDisposition != null && headers.ContentDisposition.FileName != null)
                return FileManager.RemoveInvalidFileNameChars(headers.ContentDisposition.FileName);

            var extensionFromContentType = string.Empty;
            string contentTypeName = null;
            if (headers.ContentType != null)
            {
                extensionFromContentType = MimeType.GetExtensionFromMimeType(headers.ContentType.MediaType);
                contentTypeName = headers.ContentType.Name;
            }

            if (!string.IsNullOrEmpty(headers.ContentDescription))
                return FileManager.RemoveInvalidFileNameChars(headers.ContentDescription + extensionFromContentType);

            if (!string.IsNullOrEmpty(headers.Subject))
                return FileManager.RemoveInvalidFileNameChars(headers.Subject) + extensionFromContentType;

            if (extensionFromContentType.Equals(".eml", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var message = new Message(rawBody);
                    if (message.Headers != null && !string.IsNullOrEmpty(message.Headers.Subject))
                        return FileManager.RemoveInvalidFileNameChars(message.Headers.Subject) + extensionFromContentType;
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch { }
            }

            return !string.IsNullOrEmpty(contentTypeName)
                ? FileManager.RemoveInvalidFileNameChars(contentTypeName)
                : FileManager.RemoveInvalidFileNameChars(defaultName + extensionFromContentType);
        }
        #endregion

        #region ParseBody
        /// <summary>
	    /// Parses a byte array as a body of an email message.
	    /// </summary>
	    /// <param name="rawBody">The byte array to parse as body of an email message. This array may not contain headers.</param>
	    private void ParseBody(byte[] rawBody)
	    {
	        if (IsMultiPart)
	        {
	            // Parses a MultiPart message
	            ParseMultiPartBody(rawBody);
	        }
	        else
	        {
	            // Parses a non MultiPart message
	            // Decode the body accodingly and set the Body property
	            Body = DecodeBody(rawBody, ContentTransferEncoding);
	        }
	    }
	    #endregion

        #region ParseMultiPartBody
        /// <summary>
	    /// Parses the <paramref name="rawBody"/> byte array as a MultiPart message.<br/>
	    /// It is not valid to call this method if <see cref="IsMultiPart"/> returned <see langword="false"/>.<br/>
	    /// Fills the <see cref="MessageParts"/> property of this <see cref="MessagePart"/>.
	    /// </summary>
	    /// <param name="rawBody">The byte array which is to be parsed as a MultiPart message</param>
	    private void ParseMultiPartBody(byte[] rawBody)
	    {
	        // Fetch out the boundary used to delimit the messages within the body
	        var multipartBoundary = ContentType.Boundary;

	        // Fetch the individual MultiPart message parts using the MultiPart boundary
	        var bodyParts = GetMultiPartParts(rawBody, multipartBoundary);

	        // Initialize the MessageParts property, with room to as many bodies as we have found
	        MessageParts = new List<MessagePart>(bodyParts.Count);

	        // Now parse each byte array as a message body and add it the the MessageParts property
	        foreach (var bodyPart in bodyParts)
	        {
	            var messagePart = GetMessagePart(bodyPart);
	            MessageParts.Add(messagePart);
	        }
	    }
	    #endregion

        #region GetMessagePart
        /// <summary>
	    /// Given a byte array describing a full message.<br/>
	    /// Parses the byte array into a <see cref="MessagePart"/>.
	    /// </summary>
	    /// <param name="rawMessageContent">The byte array containing both headers and body of a message</param>
	    /// <returns>A <see cref="MessagePart"/> which was described by the <paramref name="rawMessageContent"/> byte array</returns>
	    private static MessagePart GetMessagePart(byte[] rawMessageContent)
	    {
	        // Find the headers and the body parts of the byte array
	        MessageHeader headers;
	        byte[] body;
	        HeaderExtractor.ExtractHeadersAndBody(rawMessageContent, out headers, out body);

	        // Create a new MessagePart from the headers and the body
	        return new MessagePart(body, headers);
	    }
	    #endregion

        #region GetMultiPartParts
        /// <summary>
	    /// Gets a list of byte arrays where each entry in the list is a full message of a message part
	    /// </summary>
	    /// <param name="rawBody">The raw byte array describing the body of a message which is a MultiPart message</param>
	    /// <param name="multipPartBoundary">The delimiter that splits the different MultiPart bodies from each other</param>
	    /// <returns>A list of byte arrays, each a full message of a <see cref="MessagePart"/></returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="rawBody"/> is <see langword="null"/></exception>
	    private static List<byte[]> GetMultiPartParts(byte[] rawBody, string multipPartBoundary)
	    {
	        if (rawBody == null)
	            throw new ArgumentNullException("rawBody");

	        // This is the list we want to return
	        var messageBodies = new List<byte[]>();

	        // Create a stream from which we can find MultiPart boundaries
	        using (var memoryStream = new MemoryStream(rawBody))
	        {
	            bool lastMultipartBoundaryEncountered;

	            // Find the start of the first message in this multipart
	            // Since the method returns the first character on a the line containing the MultiPart boundary, we
	            // need to add the MultiPart boundary with prepended "--" and appended CRLF pair to the position returned.
	            var startLocation =
	                FindPositionOfNextMultiPartBoundary(memoryStream, multipPartBoundary,
	                    out lastMultipartBoundaryEncountered) +
	                ("--" + multipPartBoundary + "\r\n").Length;

	            while (true)
	            {
	                // When we have just parsed the last multipart entry, stop parsing on
	                if (lastMultipartBoundaryEncountered)
	                    break;

	                // Find the end location of the current multipart
	                // Since the method returns the first character on a the line containing the MultiPart boundary, we
	                // need to go a CRLF pair back, so that we do not get that into the body of the message part
	                var stopLocation =
	                    FindPositionOfNextMultiPartBoundary(memoryStream, multipPartBoundary,
	                        out lastMultipartBoundaryEncountered) -
	                    "\r\n".Length;

	                // If we could not find the next multipart boundary, but we had not yet discovered the last boundary, then
	                // we will consider the rest of the bytes as contained in a last message part.
	                if (stopLocation <= -1)
	                {
	                    // Include everything except the last CRLF.
	                    stopLocation = (int) memoryStream.Length - "\r\n".Length;

	                    // We consider this as the last part
	                    lastMultipartBoundaryEncountered = true;

	                    // Special case: when the last multipart delimiter is not ending with "--", but is indeed the last
	                    // one, then the next multipart would contain nothing, and we should not include such one.
	                    if (startLocation >= stopLocation)
	                        break;
	                }

	                // Special case: empty part.
	                // skipping by moving start location
	                if (startLocation >= stopLocation)
	                {
	                    startLocation = stopLocation + ("\r\n" + "--" + multipPartBoundary + "\r\n").Length;
	                    continue;
	                }

	                // We have now found the start and end of a message part
	                // Now we create a byte array with the correct length and put the message part's bytes into
	                // it and add it to our list we want to return
	                var length = stopLocation - startLocation;
	                var messageBody = new byte[length];
	                Array.Copy(rawBody, startLocation, messageBody, 0, length);
	                messageBodies.Add(messageBody);

	                // We want to advance to the next message parts start.
	                // We can find this by jumping forward the MultiPart boundary from the last
	                // message parts end position
	                startLocation = stopLocation + ("\r\n" + "--" + multipPartBoundary + "\r\n").Length;
	            }
	        }

	        // We are done
	        return messageBodies;
	    }
	    #endregion

        #region FindPositionOfNextMultiPartBoundary
        /// <summary>
	    /// Method that is able to find a specific MultiPart boundary in a Stream.<br/>
	    /// The Stream passed should not be used for anything else then for looking for MultiPart boundaries
	    /// <param name="stream">The stream to find the next MultiPart boundary in. Do not use it for anything else then with this method.</param>
	    /// <param name="multiPartBoundary">The MultiPart boundary to look for. This should be found in the <see cref="ContentType"/> header</param>
	    /// <param name="lastMultipartBoundaryFound">Is set to <see langword="true"/> if the next MultiPart boundary was indicated to be the last one, by having -- appended to it. Otherwise set to <see langword="false"/></param>
	    /// </summary>
	    /// <returns>The position of the first character of the line that contained MultiPartBoundary or -1 if no (more) MultiPart boundaries was found</returns>
	    private static int FindPositionOfNextMultiPartBoundary(Stream stream, string multiPartBoundary,
	        out bool lastMultipartBoundaryFound)
	    {
	        lastMultipartBoundaryFound = false;
	        while (true)
	        {
	            // Get the current position. This is the first position on the line - no characters of the line will
	            // have been read yet
	            var currentPos = (int) stream.Position;

	            // Read the line
	            var line = StreamUtility.ReadLineAsAscii(stream);

	            // If we kept reading until there was no more lines, we did not meet
	            // the MultiPart boundary. -1 is then returned to describe this.
	            if (line == null)
	                return -1;

	            // The MultiPart boundary is the MultiPartBoundary with "--" in front of it
	            // which is to be at the very start of a line
	            if (!line.StartsWith("--" + multiPartBoundary, StringComparison.Ordinal)) continue;
	            // Check if the found boundary was also the last one
	            lastMultipartBoundaryFound = line.StartsWith("--" + multiPartBoundary + "--",
	                StringComparison.OrdinalIgnoreCase);
	            return currentPos;
	        }
	    }
	    #endregion

        #region DecodeBody
        /// <summary>
		/// Decodes a byte array into another byte array based upon the Content Transfer encoding
		/// </summary>
		/// <param name="messageBody">The byte array to decode into another byte array</param>
		/// <param name="contentTransferEncoding">The <see cref="ContentTransferEncoding"/> of the byte array</param>
		/// <returns>A byte array which comes from the <paramref name="contentTransferEncoding"/> being used on the <paramref name="messageBody"/></returns>
		/// <exception cref="ArgumentNullException">If <paramref name="messageBody"/> is <see langword="null"/></exception>
		/// <exception cref="ArgumentOutOfRangeException">Thrown if the <paramref name="contentTransferEncoding"/> is unsupported</exception>
		private static byte[] DecodeBody(byte[] messageBody, ContentTransferEncoding contentTransferEncoding)
		{
			if (messageBody == null)
				throw new ArgumentNullException("messageBody");

			switch (contentTransferEncoding)
			{
				case ContentTransferEncoding.QuotedPrintable:
					// If encoded in QuotedPrintable, everything in the body is in US-ASCII
					return QuotedPrintable.DecodeContentTransferEncoding(Encoding.ASCII.GetString(messageBody));

				case ContentTransferEncoding.Base64:
					// If encoded in Base64, everything in the body is in US-ASCII
			        return Base64.Decode(Encoding.ASCII.GetString(messageBody));

				case ContentTransferEncoding.SevenBit:
				case ContentTransferEncoding.Binary:
				case ContentTransferEncoding.EightBit:
					// We do not have to do anything
					return messageBody;

				default:
					throw new ArgumentOutOfRangeException("contentTransferEncoding");
			}
		}
		#endregion

        #region GetBodyAsText
        /// <summary>
		/// Gets this MessagePart's <see cref="Body"/> as text.<br/>
		/// This is simply the <see cref="BodyEncoding"/> being used on the raw bytes of the <see cref="Body"/> property.<br/>
		/// This method is only valid to call if it is not a MultiPart message and therefore contains a body.<br/>
		/// </summary>
		/// <returns>The <see cref="Body"/> property as a string</returns>
		public string GetBodyAsText()
		{
			return BodyEncoding.GetString(Body);
		}
        #endregion

        #region Save
        /// <summary>
		/// Save this <see cref="MessagePart"/>'s contents to a file.<br/>
		/// There are no methods to reload the file.
		/// </summary>
		/// <param name="file">The File location to save the <see cref="MessagePart"/> to. Existent files will be overwritten.</param>
		/// <exception cref="ArgumentNullException">If <paramref name="file"/> is <see langword="null"/></exception>
		/// <exception>Other exceptions relevant to using a <see cref="FileStream"/> might be thrown as well</exception>
		public void Save(FileInfo file)
		{
			if (file == null)
				throw new ArgumentNullException("file");

			using (var fileStream = new FileStream(file.FullName, FileMode.Create))
			{
				Save(fileStream);
			}
		}

		/// <summary>
		/// Save this <see cref="MessagePart"/>'s contents to a stream.<br/>
		/// </summary>
		/// <param name="messageStream">The stream to write to</param>
		/// <exception cref="ArgumentNullException">If <paramref name="messageStream"/> is <see langword="null"/></exception>
		/// <exception>Other exceptions relevant to <see cref="Stream.Write"/> might be thrown as well</exception>
		public void Save(Stream messageStream)
		{
			if (messageStream == null)
				throw new ArgumentNullException("messageStream");

			messageStream.Write(Body, 0, Body.Length);
		}
		#endregion
	}
}