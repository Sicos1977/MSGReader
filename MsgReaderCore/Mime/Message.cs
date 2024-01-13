using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using MsgReader.Helpers;
using MsgReader.Mime.Header;
using MsgReader.Mime.Traverse;

// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable UnusedMember.Global
// ReSharper disable IntroduceOptionalParameters.Global

namespace MsgReader.Mime;

/// <summary>
///     This is the root of the email tree structure.<br />
///     <see cref="Mime.MessagePart" /> for a description about the structure.<br />
///     <br />
///     A Message (this class) contains the headers of an email message such as:
///     <code>
///  - To
///  - From
///  - Subject
///  - Content-Type
///  - Message-ID
/// </code>
///     which are located in the <see cref="Headers" /> property.<br />
///     <br />
///     Use the <see cref="Message.MessagePart" /> property to find the actual content of the email message.
/// </summary>
public class Message
{
    #region Fields
    private bool _changed;
    #endregion

    #region Properties
    /// <summary>
    ///     Returns the ID of the message when this is available in the <see cref="Headers" />
    ///     (as specified in [RFC2822]). Null when not available
    /// </summary>
    public string Id
    {
        get
        {
            if (Headers == null)
                return null;

            return !string.IsNullOrEmpty(Headers.MessageId) ? Headers.MessageId : null;
        }
    }

    /// <summary>
    ///     Headers of the Message.
    /// </summary>
    public MessageHeader Headers { get; private set; }

    /// <summary>
    ///     This is the body of the email Message.<br />
    ///     <br />
    ///     If the body was parsed for this Message, this property will never be <see langword="null" />.
    /// </summary>
    public MessagePart MessagePart { get; private set; }

    /// <summary>
    ///     This will return the first <see cref="MessagePart" /> where the <see cref="ContentType.MediaType" />
    ///     is set to "html/text". This will return <see langword="null" /> when there is no "html/text"
    ///     <see cref="MessagePart" /> found.
    /// </summary>
    public MessagePart HtmlBody { get; private set; }

    /// <summary>
    ///     This will return the first <see cref="MessagePart" /> where the <see cref="ContentType.MediaType" />
    ///     is set to "text/plain". This will be <see langword="null" /> when there is no "text/plain"
    ///     <see cref="MessagePart" /> found.
    /// </summary>
    public MessagePart TextBody { get; private set; }

    /// <summary>
    ///     This will return all the <see cref="MessagePart">message parts</see> that are flagged as
    ///     <see cref="Mime.MessagePart.IsAttachment" />.
    ///     This will be <see langword="null" /> when there are no <see cref="MessagePart">message parts</see>
    ///     that are flagged as <see cref="Mime.MessagePart.IsAttachment" />.
    /// </summary>
    public ObservableCollection<MessagePart> Attachments { get; private set; }

    /// <summary>
    ///     The raw content from which this message has been constructed.<br />
    ///     These bytes can be persisted and later used to recreate the Message.
    /// </summary>
    public byte[] RawMessage { get; private set; }

    /// <summary>
    ///     Returns <c>true</c> when the signature is valid />
    /// </summary>
    public bool? SignatureIsValid { get; private set; }

    /// <summary>
    ///     Returns the name of the person who signed the message
    /// </summary>
    public string SignedBy { get; private set; }

    /// <summary>
    ///     Returns the date and time when the message has been signed
    /// </summary>
    public DateTime? SignedOn { get; private set; }

    /// <summary>
    ///     Returns the certificate that has been used to sign the message
    /// </summary>
    public X509Certificate2 SignedCertificate { get; private set; }
    #endregion

    #region Constructors
    /// <summary>
    ///     Convenience constructor for <see cref="Mime.Message(byte[], bool)" />.<br />
    ///     <br />
    ///     Creates a message from a byte array. The full message including its body is parsed.
    /// </summary>
    /// <param name="rawMessageContent">The byte array which is the message contents to parse</param>
    public Message(byte[] rawMessageContent) : this(rawMessageContent, true)
    {
    }

    /// <summary>
    ///     Convenience constructor for <see cref="Mime.Message(byte[], bool)" />.<br />
    ///     <br />
    ///     Creates a message from a stream. The full message including its body is parsed.
    /// </summary>
    /// <param name="rawMessageContent">The byte array which is the message contents to parse</param>
    public Message(Stream rawMessageContent)
    {
        using var recyclableMemoryStream = StreamHelpers.Manager.GetStream("Message.cs");
        rawMessageContent.CopyTo(recyclableMemoryStream);
        ParseContent(recyclableMemoryStream.ToArray(), true);
    }

    /// <summary>
    ///     Constructs a message from a byte array.<br />
    ///     <br />
    ///     The headers are always parsed, but if <paramref name="parseBody" /> is <see langword="false" />, the body is not
    ///     parsed.
    /// </summary>
    /// <param name="rawMessageContent">The byte array which is the message contents to parse</param>
    /// <param name="parseBody">
    ///     <see langword="true" /> if the body should be parsed,
    ///     <see langword="false" /> if only headers should be parsed out of the <paramref name="rawMessageContent" /> byte
    ///     array
    /// </param>
    public Message(byte[] rawMessageContent, bool parseBody)
    {
        ParseContent(rawMessageContent, parseBody);
    }
    #endregion

    #region ParseContent
    /// <summary>
    ///     Constructs a message from a byte array.<br />
    ///     <br />
    ///     The headers are always parsed, but if <paramref name="parseBody" /> is <see langword="false" />, the body is not
    ///     parsed.
    /// </summary>
    /// <param name="rawMessageContent">The byte array which is the message contents to parse</param>
    /// <param name="parseBody">
    ///     <see langword="true" /> if the body should be parsed,
    ///     <see langword="false" /> if only headers should be parsed out of the <paramref name="rawMessageContent" /> byte
    ///     array
    /// </param>
    private void ParseContent(byte[] rawMessageContent, bool parseBody)
    {
        Logger.WriteToLog("Processing raw EML message content");

        RawMessage = rawMessageContent;

        // Find the headers and the body parts of the byte array
        HeaderExtractor.ExtractHeadersAndBody(rawMessageContent, out var headersTemp, out var body);

        // Set the Headers property
        Headers = headersTemp;

        // Should we also parse the body?
        if (parseBody)
        {
            // Parse the body into a MessagePart
            MessagePart = new MessagePart(body, Headers);

            var attachments = new AttachmentFinder().VisitMessage(this);

            if (MessagePart.ContentType?.MediaType == "multipart/signed")
                foreach (var attachment in attachments)
                    if (attachment.FileName.ToUpperInvariant() == "SMIME.P7S")
                        ParseContent(ProcessSignedContent(attachment.Body), true);

            var findBodyMessagePartWithMediaType = new FindBodyMessagePartWithMediaType();

            // Searches for the first HTML body and mark this one as the HTML body of the E-mail
            if (HtmlBody == null)
            {
                Logger.WriteToLog("There was not HTML body found, trying to find one");

                var index = attachments.FindIndex(m => m.IsHtmlBody);
                if (index != -1)
                {
                    Logger.WriteToLog("Found HTML attachment setting it as the HTML body");
                    HtmlBody = attachments[index];
                    attachments.RemoveAt(index);
                }
                else
                {
                    HtmlBody = findBodyMessagePartWithMediaType.VisitMessage(this, "text/html");
                    if (HtmlBody != null)
                    {
                        Logger.WriteToLog("Found HTML message part setting it as the HTML body");
                        HtmlBody.IsHtmlBody = true;
                    }
                    else
                    {
                        index = attachments.FindIndex(m => m.ContentType?.MediaType == "text/html");
                        if (index != -1)
                        {
                            Logger.WriteToLog("Found HTML attachment setting it as the HTML body");
                            HtmlBody = attachments[index];
                            attachments.RemoveAt(index);
                        }
                    }
                }
            }

            // Searches for the first TEXT body and mark this one as the TEXT body of the E-mail
            if (TextBody == null)
            {
                Logger.WriteToLog("There was not TEXT body found, trying to find one");

                var index = attachments.FindIndex(m => m.IsTextBody);
                if (index != -1)
                {
                    Logger.WriteToLog("Found TEXT attachment setting it as the TEXT body");
                    TextBody = attachments[index];
                    attachments.RemoveAt(index);
                }
                else
                {
                    TextBody = findBodyMessagePartWithMediaType.VisitMessage(this, "text/plain");
                    if (TextBody != null)
                    {
                        Logger.WriteToLog("Found TEXT message part setting it as the TEXT body");
                        TextBody.IsTextBody = true;
                    }
                    else
                    {
                        index = attachments.FindIndex(m => m.ContentType?.MediaType == "text/plain");
                        if (index != -1)
                        {
                            Logger.WriteToLog("Found TEXT attachment setting it as the TEXT body");
                            HtmlBody = attachments[index];
                            attachments.RemoveAt(index);
                        }
                    }
                }
            }

            if (HtmlBody != null)
                foreach (var attachment in attachments)
                {
                    if (attachment.IsInline || attachment.ContentId == null ||
                        attachment.FileName.ToUpperInvariant() == "SMIME.P7S")
                        continue;

                    var htmlBody = HtmlBody.BodyEncoding.GetString(HtmlBody.Body);
                    attachment.IsInline = htmlBody.Contains($"cid:{attachment.ContentId}");
                }

            if (attachments != null)
            {
                var result = new ObservableCollection<MessagePart>();
                foreach (var attachment in attachments)
                {
                    if (attachment.FileName.ToUpperInvariant() == "SMIME.P7S" ||
                        attachment.IsHtmlBody ||
                        attachment.IsTextBody)
                        continue;

                    result.Add(attachment);
                }

                Attachments = result;
                Attachments.CollectionChanged += (_, _) => { _changed = true; };
            }
        }

        Logger.WriteToLog("Raw EML message content processed");
    }
    #endregion

    #region ProcessSignedContent
    /// <summary>
    ///     Processes the signed content
    /// </summary>
    /// <param name="data"></param>
    /// <returns></returns>
    private byte[] ProcessSignedContent(byte[] data)
    {
        Logger.WriteToLog("Processing signed content");

        var signedCms = new SignedCms();
        signedCms.Decode(data);

        try
        {
            SignatureIsValid = true;

            //signedCms.CheckSignature(signedCms.Certificates, false);
            foreach (var cert in signedCms.Certificates)
                if (!cert.Verify())
                    SignatureIsValid = false;

            foreach (var cryptographicAttributeObject in signedCms.SignerInfos[0].SignedAttributes)
                if (cryptographicAttributeObject.Values[0] is Pkcs9SigningTime pkcs9SigningTime)
                    SignedOn = pkcs9SigningTime.SigningTime.ToLocalTime();

            var certificate = signedCms.SignerInfos[0].Certificate;
            if (certificate != null)
            {
                SignedCertificate = certificate;
                SignedBy = certificate.GetNameInfo(X509NameType.SimpleName, false);
            }
        }
        catch (CryptographicException)
        {
            SignatureIsValid = false;
        }

        // Get the decoded attachment
        using var recyclableMemoryStream = StreamHelpers.Manager.GetStream("Message.cs", signedCms.ContentInfo.Content, 0, signedCms.ContentInfo.Content.Length);
        Logger.WriteToLog("Signed content processed");
        return recyclableMemoryStream.ToArray();
    }
    #endregion

    #region GetEmailAddresses
    /// <summary>
    ///     Returns the list of <see cref="RfcMailAddress" /> as a normal or html string
    /// </summary>
    /// <param name="rfcMailAddresses">A list with one or more <see cref="RfcMailAddress" /> objects</param>
    /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
    /// <param name="html">Set this to true when the E-mail body format is html</param>
    /// <returns></returns>
    public string GetEmailAddresses(IEnumerable<RfcMailAddress> rfcMailAddresses, bool convertToHref, bool html)
    {
        Logger.WriteToLog("Getting mail addresses");

        var result = string.Empty;

        if (rfcMailAddresses == null)
            return result;

        foreach (var rfcMailAddress in rfcMailAddresses)
        {
            if (result != string.Empty)
                result += "; ";

            var emailAddress = string.Empty;
            var displayName = rfcMailAddress.DisplayName;

            if (rfcMailAddress.HasValidMailAddress)
                emailAddress = rfcMailAddress.Address;

            if (string.Equals(emailAddress, displayName, StringComparison.InvariantCultureIgnoreCase))
                displayName = string.Empty;

            if (html)
            {
                emailAddress = WebUtility.HtmlEncode(emailAddress);
                displayName = WebUtility.HtmlEncode(displayName);
            }

            if (convertToHref && html && !string.IsNullOrEmpty(emailAddress))
            {
                result += "<a href=\"mailto:" + emailAddress + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                              ? displayName
                              : emailAddress) + "</a>";
            }

            else
            {
                if (!string.IsNullOrEmpty(displayName))
                    result += displayName;

                var beginTag = string.Empty;
                var endTag = string.Empty;
                if (!string.IsNullOrEmpty(displayName))
                {
                    if (html)
                    {
                        beginTag = "&nbsp;&lt;";
                        endTag = "&gt;";
                    }
                    else
                    {
                        beginTag = " <";
                        endTag = ">";
                    }
                }

                if (!string.IsNullOrEmpty(emailAddress))
                    result += beginTag + emailAddress + endTag;
            }
        }

        return result;
    }
    #endregion

    #region ToMailMessage
    /// <summary>
    /// This method will convert this <see cref="Message"/> into a <see cref="MailMessage"/> equivalent.<br/>
    /// The returned <see cref="MailMessage"/> can be used with <see cref="System.Net.Mail.SmtpClient"/> to forward the email.<br/>
    /// <br/>
    /// You should be aware of the following about this method:
    /// <list type="bullet">
    /// <item>
    ///    All sender and receiver mail addresses are set.
    ///    If you send this email using a <see cref="System.Net.Mail.SmtpClient"/> then all
    ///    receivers in To, From, Cc and Bcc will receive the email once again.
    /// </item>
    /// <item>
    ///    If you view the source code of this Message and looks at the source code of the forwarded
    ///    <see cref="MailMessage"/> returned by this method, you will notice that the source codes are not the same.
    ///    The content that is presented by a mail client reading the forwarded <see cref="MailMessage"/> should be the
    ///    same as the original, though.
    /// </item>
    /// <item>
    ///    Content-Disposition headers will not be copied to the <see cref="MailMessage"/>.
    ///    It is simply not possible to set these on Attachments.
    /// </item>
    /// <item>
    ///    HTML content will be treated as the preferred view for the <see cref="MailMessage.Body"/>. Plain text content will be used for the
    ///    <see cref="MailMessage.Body"/> when HTML is not available.
    /// </item>
    /// </list>
    /// </summary>
    /// <returns>A <see cref="MailMessage"/> object that contains the same information that this Message does</returns>
    private MailMessage ToMailMessage()
    {
        // Construct an empty MailMessage to which we will gradually build up to look like the current Message object (this)
        var message = new MailMessage { Subject = Headers.Subject, SubjectEncoding = Encoding.UTF8 };

        AlternateView htmlView = null;

        if (HtmlBody != null)
        {
            htmlView = AlternateView.CreateAlternateViewFromString(HtmlBody.GetBodyAsText(), Encoding.UTF8, "text/html");
            message.AlternateViews.Add(htmlView);
        }

        if (TextBody != null)
            message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(TextBody.GetBodyAsText(), Encoding.UTF8, "text/plain"));
        
        // Add attachments to the message
        foreach (var attachmentMessagePart in Attachments)
        {
            var recyclableMemoryStream = StreamHelpers.Manager.GetStream("Message.cs", attachmentMessagePart.Body, 0, attachmentMessagePart.Body.Length);

            if (attachmentMessagePart.IsInline && htmlView != null)
            {
                var linkedResource = new LinkedResource(recyclableMemoryStream, attachmentMessagePart.ContentType)
                {
                    ContentId = attachmentMessagePart.ContentId
                };

                htmlView.LinkedResources.Add(linkedResource);
            }
            else
            {
                var attachment = new Attachment(recyclableMemoryStream, attachmentMessagePart.ContentType)
                {
                    ContentId = attachmentMessagePart.ContentId,
                    Name = attachmentMessagePart.FileName
                };

                message.Attachments.Add(attachment);
            }
        }

        if (Headers.From is { HasValidMailAddress: true })
            message.From = Headers.From.MailAddress;

        if (Headers.ReplyTo is { HasValidMailAddress: true })
            message.ReplyToList.Add(Headers.ReplyTo.MailAddress);

        if (Headers.Sender is { HasValidMailAddress: true })
            message.Sender = Headers.Sender.MailAddress;

        foreach (var to in Headers.To.Where(to => to.HasValidMailAddress))
            message.To.Add(to.MailAddress);

        foreach (var cc in Headers.Cc.Where(cc => cc.HasValidMailAddress))
            message.CC.Add(cc.MailAddress);

        foreach (var bcc in Headers.Bcc.Where(bcc => bcc.HasValidMailAddress))
            message.Bcc.Add(bcc.MailAddress);

        return message;
    }
    #endregion

    #region Save
    /// <summary>
    ///     Save this <see cref="Message" /> to a file.<br />
    ///     <br />
    /// </summary>
    /// <param name="file">The File location to save the <see cref="Message" /> to. Existent files will be overwritten.</param>
    /// <exception cref="ArgumentNullException">If <paramref name="file" /> is <see langword="null" /></exception>
    /// <exception>Other exceptions relevant to using a <see cref="FileStream" /> might be thrown as well</exception>
    // ReSharper disable once UnusedMember.Global
    public void Save(FileInfo file)
    {
        if (file == null)
            throw new ArgumentNullException(nameof(file));

        Save(file.OpenWrite());
    }

    /// <summary>
    ///     Save this <see cref="Message" /> to a stream.<br />
    /// </summary>
    /// <param name="messageStream">The stream to write to</param>
    /// <exception cref="ArgumentNullException">If <paramref name="messageStream" /> is <see langword="null" /></exception>
    /// <exception>Other exceptions relevant to Stream.Write might be thrown as well</exception>
    public void Save(Stream messageStream)
    {
        if (messageStream == null)
            throw new ArgumentNullException(nameof(messageStream));

        if (_changed)
        {
            ToMailMessage().WriteTo(messageStream);
            Logger.WriteToLog("EML message saved as new message");
        }
        else
        {
            messageStream.Write(RawMessage, 0, RawMessage.Length);
            Logger.WriteToLog("Raw EML message saved");
        }
    }
    #endregion

    #region Load
    /// <summary>
    ///     Loads a <see cref="Message" /> from a file containing a raw email.
    /// </summary>
    /// <param name="file">The File location to load the <see cref="Message" /> from. The file must exist.</param>
    /// <exception cref="ArgumentNullException">If <paramref name="file" /> is <see langword="null" /></exception>
    /// <exception cref="FileNotFoundException">If <paramref name="file" /> does not exist</exception>
    /// <exception>Other exceptions relevant to a <see cref="FileStream" /> might be thrown as well</exception>
    /// <returns>A <see cref="Message" /> with the content loaded from the <paramref name="file" /></returns>
    public static Message Load(FileInfo file)
    {
        Logger.WriteToLog($"Loading EML file from '{file.FullName}'");

        if (file == null)
            throw new ArgumentNullException(nameof(file));

        if (!file.Exists)
            throw new FileNotFoundException("Cannot load message from non-existent file", file.FullName);

        return Load(file.OpenRead());
    }

    /// <summary>
    ///     Loads a <see cref="Message" /> from a <see cref="Stream" /> containing a raw email.
    /// </summary>
    /// <param name="messageStream">The <see cref="Stream" /> from which to load the raw <see cref="Message" /></param>
    /// <exception cref="ArgumentNullException">If <paramref name="messageStream" /> is <see langword="null" /></exception>
    /// <exception>Other exceptions relevant to Stream.Read might be thrown as well</exception>
    /// <returns>A <see cref="Message" /> with the content loaded from the <paramref name="messageStream" /></returns>
    public static Message Load(Stream messageStream)
    {
        Logger.WriteToLog("Loading EML file from stream");

        if (messageStream == null)
            throw new ArgumentNullException(nameof(messageStream));

        using var recyclableMemoryStream = StreamHelpers.Manager.GetStream();
        messageStream.CopyTo(recyclableMemoryStream);
        var content = recyclableMemoryStream.ToArray();
        return new Message(content);
    }
    #endregion
}