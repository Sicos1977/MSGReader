//
// Reader.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2024 Kees van Spelde. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using MsgReader.Exceptions;
using MsgReader.Helpers;
using MsgReader.Localization;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

#if (NETFRAMEWORK)
using System.Drawing;
using System.Drawing.Imaging;
#endif

namespace MsgReader
{
    #region Interface IReader
    /// <summary>
    /// Interface to make Reader class COM visible
    /// </summary>
    public interface IReader
    {
        /// <summary>
        /// Extract the input msg file to the given output folder
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to extract the msg file</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <param name="culture">Sets the culture that needs to be used to localize the output of this class</param>
        /// <returns>String array containing the message body and its (inline) attachments</returns>
        [DispId(1)]
        // ReSharper disable once UnusedMemberInSuper.Global
        string[] ExtractToFolderFromCom(string inputFile, string outputFolder, ReaderHyperLinks hyperlinks = ReaderHyperLinks.None, string culture = "");

        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        // ReSharper disable once UnusedMemberInSuper.Global
        // ReSharper disable once UnusedMember.Global
        string GetErrorMessage();
    }
    #endregion

    #region ReaderHyperLinks
    /// <summary>
    /// Tells the readers class when the generate hyperlinks
    /// </summary>
    public enum ReaderHyperLinks
    {
        /// <summary>
        /// Do not generate any hyperlink
        /// </summary>
        None,

        /// <summary>
        /// Only generate hyperlinks for the e-mail addresses
        /// </summary>
        Email,

        /// <summary>
        /// Only generate hyperlinks for the attachments
        /// </summary>
        Attachments,

        /// <summary>
        /// Generate hyperlinks for the e-mail addresses and attachments
        /// </summary>
        Both
    }
    #endregion

    /// <summary>
    /// This class can be used to read an Outlook msg file and save the message body (in HTML or TEXT format)
    /// and all it's attachments to an output folder.
    /// </summary>
    [Guid("E9641DF0-18FC-11E2-BC95-1ACF6088709B")]
    [ComVisible(true)]
    public class Reader : IReader
    {
        #region Fields
        /// <summary>
        /// Contains an error message when something goes wrong in the <see cref="ExtractToFolderFromCom"/> method.
        /// This message can be retrieved with the GetErrorMessage. This way we keep .NET exceptions inside
        /// when this code is called from a COM language
        /// </summary>
        private string _errorMessage;

        /// <summary>
        /// Used to keep track if we already did write an empty line
        /// </summary>
        private static bool _emptyLineWritten;

        /// <summary>
        /// Placeholder for custom header styling
        /// </summary>
        private static string _customHeaderStyleCss;
        #endregion

        #region Properties
        /// <summary>
        ///     A unique id that can be used to identify the logging of the reader when
        ///     calling the code from multiple threads and writing all the logging to the same file
        /// </summary>
        public string InstanceId
        {
            set => Logger.InstanceId = value;
        }

        /// <summary>
        /// Set / Get whether to use default styling of email header or
        /// to use the custom CSS style set by <see cref="SetCustomHeaderStyle"/>
        /// </summary>
        public bool UseCustomHeaderStyle
        {
            get;
            set;
        }

        /// <summary>
        /// If true the header is injected as an iframe effectively ensuring it is not affected by any css in the message
        /// </summary>
        public bool InjectHeaderAsIFrame
        {
            get;
            set;
        }

        /// <summary>
        ///     When an MSG file contains an RTF file with encapsulated HTML and the RTF
        ///     uses fonts with different encodings then this levels set the threshold that
        ///     an encoded string detection levels needs to be before recognizing it as a valid
        ///     string. When the detection level is lower than this setting then the default RTF
        ///     encoding is used to decode the encoded char 
        /// </summary>
        /// <remarks>
        ///     Default this value is set to 0.90, any values lower than 0.70 probably give bad
        ///     results
        /// </remarks>
        public float CharsetDetectionEncodingConfidenceLevel { get; set; } = 0.90f;
        #endregion

        #region HeaderStyle
        /// <summary>
        /// Set the custom CSS stylesheet for the email header.
        /// Set to string.Empty or null to reset to default. Get current or default CSS via <see cref="GetCustomHeaderStyle"/>
        /// </summary>
        /// <param name="headerStyleCss"></param>
        public static void SetCustomHeaderStyle(string headerStyleCss)
        {
            _customHeaderStyleCss = headerStyleCss;
        }

        /// <summary>
        /// Get current custom CSS stylesheet to apply to email header
        /// </summary>
        /// <returns>Returns default CSS until a custom is set via <see cref="SetCustomHeaderStyle"/></returns>
        public static string GetCustomHeaderStyle()
        {
            if (!string.IsNullOrEmpty(_customHeaderStyleCss))
                return _customHeaderStyleCss;

            // Return defaultStyle
            const string defaultHeaderCss =
                "table.MsgReaderHeader {" +
                "   font-family: Times New Roman; font-size: 12pt;" +
                "}\n" +
                "tr.MsgReaderHeaderRow {" +
                "   height: 18px; vertical-align: top;" +
                "}\n" +
                "tr.MsgReaderHeaderRowEmpty {}\n" +
                "td.MsgReaderHeaderRowLabel {" +
                "   font-weight: bold; white-space:nowrap;" +
                "}\n" +
                "td.MsgReaderHeaderRowText {}\n" +
                "div.MsgReaderContactPhoto {" +
                "   height: 250px; position: absolute; top: 20px; right: 20px;" +
                "}\n" +
                "div.MsgReaderContactPhoto > img {" +
                "   height: 100%;" +
                "}\n" +
                "table.MsgReaderInlineAttachment {" +
                "   width: 70px; display: inline; text-align: center; font-family: Times New Roman; font-size: 12pt;" +
                "}";

            return defaultHeaderCss;
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets it's needed properties
        /// </summary>
        /// <param name="logStream">When set then logging is written to this stream for all extractions. If
        /// you want a separate log for each extraction then set the log stream on one of the ExtractTo methods</param>
        public Reader(Stream logStream = null)
        {
#if NET5_0_OR_GREATER
            var encodingProvider = CodePagesEncodingProvider.Instance;
            Encoding.RegisterProvider(encodingProvider);
#endif

            if (logStream != null)
                Logger.LogStream = logStream;
        }
        #endregion

        #region SetCulture
        /// <summary>
        /// Sets the culture that needs to be used to localize the output of this class. 
        /// Default the current system culture is set. When there is no localization available the
        /// default will be used. This will be en-US.
        /// </summary>
        /// <param name="name">The name of the culture e.g. nl-NL</param>
        public void SetCulture(string name)
        {
            Logger.WriteToLog($"Setting culture to '{name}'");
            Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo(name);
        }
        #endregion

        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <paramref name="inputFile"/> and <paramref name="outputFolder"/> is valid
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFolder"></param>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <paramref name="outputFolder"/> does not exist</exception>
        /// <exception cref="MRFileTypeNotSupported">Raised when the extension is not .msg or .eml</exception>
        private static string CheckFileNameAndOutputFolder(string inputFile, string outputFolder)
        {
            Logger.WriteToLog("Checking input file and output folder");

            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException(inputFile);

            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentNullException(outputFolder);

            if (!File.Exists(inputFile))
                throw new FileNotFoundException(inputFile);

            if (!Directory.Exists(outputFolder))
                throw new DirectoryNotFoundException(outputFolder);

            var extension = Path.GetExtension(inputFile);
            if (string.IsNullOrEmpty(extension))
                throw new MRFileTypeNotSupported("Expected .msg or .eml extension on the input file");

            extension = extension.ToUpperInvariant();

            using var fileStream = File.OpenRead(inputFile);
            var header = new byte[2];
            // ReSharper disable once MustUseReturnValue
            fileStream.Read(header, 0, 2);

            switch (extension)
            {
                case ".MSG":
                case ".EML":

                    // Sometimes the email contains an MSG extension and actual it's an EML.
                    // Very often this happens when a user saves the email manually and types 
                    // the filename. To prevent these kind of errors we do a double check to make sure 
                    // the file is really an MSG file
                    if (header[0] == 0xD0 && header[1] == 0xCF)
                        return ".MSG";

                    return ".EML";

                default:
                    const string message = "Wrong file extension, expected .msg or .eml";
                    throw new MRFileTypeNotSupported(message);
            }
        }
        #endregion

        #region ExtractToStream
        /// <summary>
        /// This method reads the <paramref name="inputStream"/> and when the stream is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body, so it looks
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachments <br/>
        /// And in the end returns everything to the output stream
        /// </summary>
        /// <param name="inputStream">The mime stream</param>
        /// <param name="hyperlinks">When true hyperlinks are generated for the To, CC, BCC and attachments</param>
        public List<MemoryStream> ExtractToStream(MemoryStream inputStream, bool hyperlinks = false)
        {
            var message = Mime.Message.Load(inputStream);
            return WriteEmlStreamEmail(message, hyperlinks);
        }
        #endregion

        #region ExtractToFolder
        /// <summary>
        /// This method reads the <paramref name="inputFile"/> and when the file is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body, so it looks 
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachments <br/>
        /// And in the end writes everything to the given <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to save the extracted msg file</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <param name="culture"></param>
        public string[] ExtractToFolderFromCom(string inputFile,
            string outputFolder,
            ReaderHyperLinks hyperlinks = ReaderHyperLinks.None,
            string culture = "")
        {
            try
            {
                if (!string.IsNullOrEmpty(culture))
                    SetCulture(culture);

                return ExtractToFolder(inputFile, outputFolder, hyperlinks);
            }
            catch (Exception e)
            {
                _errorMessage = ExceptionHelpers.GetInnerException(e);
                return Array.Empty<string>();
            }
        }

        /// <summary>
        /// This method reads the <paramref name="inputFile"/> and when the file is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body, so it looks 
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachments <br/>
        /// And in the end writes everything to the given <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to save the extracted msg file</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <param name="messageType">Use this if you get the exception <see cref="MRFileTypeNotSupported"/> and
        /// want to force this method to use a specific <see cref="MessageType"/> to parse this MSG file. This
        /// is only used when the file is an MSG file</param>
        /// <param name="logStream">When set then this will give a logging for each extraction. Use the log stream
        /// option in the constructor if you want one log for all extractions</param>
        /// <param name="includeReactionsInfo">When <c>true</c> then reactions information is also included in the output</param>
        /// <returns>String array containing the full path to the message body and its attachments</returns>
        /// <exception cref="MRFileTypeNotSupported">Raised when the Microsoft Outlook message type is not supported</exception>
        /// <exception cref="MRInvalidSignedFile">Raised when the Microsoft Outlook signed message is invalid</exception>
        /// <exception cref="ArgumentNullException">Raised when the <param ref="inputFile"/> or <param ref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <param ref="inputFile"/> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <param ref="outputFolder"/> does not exist</exception>
        public string[] ExtractToFolder(
            string inputFile,
            string outputFolder,
            ReaderHyperLinks hyperlinks = ReaderHyperLinks.None,
            MessageType? messageType = null,
            Stream logStream = null,
            bool includeReactionsInfo = false)
        {
            if (logStream != null)
                Logger.LogStream = logStream;

            outputFolder = FileManager.CheckForDirectorySeparator(outputFolder);

            _errorMessage = string.Empty;

            var extension = CheckFileNameAndOutputFolder(inputFile, outputFolder);

            switch (extension)
            {
                case ".EML":
                    Logger.WriteToLog($"Extracting EML file '{inputFile}' to output folder '{outputFolder}'");

                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        var message = Mime.Message.Load(stream);
                        return WriteEmlEmail(message, outputFolder, hyperlinks).ToArray();
                    }

                case ".MSG":
                    Logger.WriteToLog($"Extracting MSG file '{inputFile}' to output folder '{outputFolder}'");

                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    using (var message = new Storage.Message(stream))
                    {
                        messageType ??= message.Type;

                        Logger.WriteToLog($"MSG file has the type '{messageType}'");

                        switch (messageType)
                        {
                            case MessageType.Email:
                            case MessageType.EmailSms:
                            case MessageType.EmailTemplateMicrosoft:
                            case MessageType.EmailNonDeliveryReport:
                            case MessageType.EmailDeliveryReport:
                            case MessageType.EmailDelayedDeliveryReport:
                            case MessageType.EmailReadReceipt:
                            case MessageType.EmailNonReadReceipt:
                            case MessageType.EmailEncryptedAndMaybeSigned:
                            case MessageType.EmailEncryptedAndMaybeSignedNonDelivery:
                            case MessageType.EmailEncryptedAndMaybeSignedDelivery:
                            case MessageType.EmailClearSignedReadReceipt:
                            case MessageType.EmailClearSignedNonDelivery:
                            case MessageType.EmailClearSignedDelivery:
                            case MessageType.EmailBmaStub:
                            case MessageType.CiscoUnityVoiceMessage:
                            case MessageType.EmailClearSigned:
                            case MessageType.RightFaxAdv:
                            case MessageType.SkypeForBusinessMissedMessage:
                            case MessageType.SkypeForBusinessConversation:
                            case MessageType.WorkSiteEmsFiled:
                            case MessageType.WorkSiteEmsFiledRe:
                            case MessageType.WorkSiteEmsQueued:
                            case MessageType.WorkSiteEmsError:
                            case MessageType.WorkSiteEmsFiledFw:
                            case MessageType.WorkSiteEmsFw:
                            case MessageType.WorkSiteEmsRe:
                            case MessageType.WorkSiteEmsSent:
                            case MessageType.WorkSiteEmsSentFw:
                            case MessageType.WorkSiteEmsSentRe:
                            case MessageType.SkypeTeamsMessage:
                                return WriteMsgEmail(message, outputFolder, hyperlinks, includeReactionsInfo).ToArray();

                            case MessageType.Appointment:
                            case MessageType.AppointmentNotification:
                            case MessageType.AppointmentSchedule:
                            case MessageType.AppointmentRequest:
                            case MessageType.AppointmentRequestNonDelivery:
                            case MessageType.AppointmentResponse:
                            case MessageType.AppointmentResponseCanceled:
                            case MessageType.AppointmentResponseCanceledNonDelivery:
                            case MessageType.AppointmentResponsePositive:
                            case MessageType.AppointmentResponsePositiveNonDelivery:
                            case MessageType.AppointmentResponseNegative:
                            case MessageType.AppointmentResponseNegativeNonDelivery:
                            case MessageType.AppointmentResponseTentative:
                            case MessageType.AppointmentResponseTentativeNonDelivery:
                                return WriteMsgAppointment(message, outputFolder, hyperlinks).ToArray();

                            case MessageType.Contact:
                                return WriteMsgContact(message, outputFolder, hyperlinks).ToArray();

                            case MessageType.Task:
                            case MessageType.TaskRequestAccept:
                            case MessageType.TaskRequestDecline:
                            case MessageType.TaskRequestUpdate:
                                return WriteMsgTask(message, outputFolder, hyperlinks).ToArray();

                            case MessageType.StickyNote:
                                return WriteMsgStickyNote(message, outputFolder).ToArray();

                            case MessageType.Journal:
                                return WriteMsgJournal(message, outputFolder, hyperlinks).ToArray();


                            case MessageType.Unknown:
                                const string unknown = "Unsupported message type";
                                Logger.WriteToLog(unknown);
                                throw new MRFileTypeNotSupported(unknown);

                            default:
                                throw new ArgumentOutOfRangeException(nameof(messageType), messageType, null);
                        }
                    }
            }

            return Array.Empty<string>();
        }
        #endregion

        #region ExtractMessageBody
        /// <summary>
        /// Extract a mail body in memory without saving data on the hard drive.
        /// </summary>
        /// <param name="stream">The message as a stream</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <param name="contentType">Content type, e.g. text/html; charset=utf-8</param>
        /// <param name="withHeaderTable">
        /// When true, a text/html table with information of To, CC, BCC and attachments will
        /// be generated and inserted at the top of the text/html document
        /// </param>
        /// <param name="includeReactionsInfo">When <c>true</c> then reactions information is also included in the output</param>
        /// <returns>Body as string (can be html code, ...)</returns>
        public string ExtractMsgEmailBody(Stream stream, ReaderHyperLinks hyperlinks, string contentType, bool withHeaderTable = true, bool includeReactionsInfo = false)
        {
            Logger.WriteToLog("Extracting EML message body from a stream");

            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            // Reset stream to be sure we start at the beginning
            stream.Seek(0, SeekOrigin.Begin);

            using var message = new Storage.Message(stream);
            return ExtractMsgEmailBody(message, hyperlinks, contentType, withHeaderTable, includeReactionsInfo: includeReactionsInfo);
        }

        /// <summary>
        /// Extract a mail body in memory without saving data on the hard drive.
        /// </summary>
        /// <param name="message">The message as a stream</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <param name="contentType">Content type, e.g. text/html; charset=utf-8</param>
        /// <param name="withHeaderTable">
        /// When true, a text/html table with information of To, CC, BCC and attachments will
        /// be generated and inserted at the top of the text/html document
        /// </param>
        /// <param name="includeReactionsInfo">When <c>true</c> then reactions information is also included in the output</param>
        /// <returns>Body as string (can be html code, ...)</returns>
        public string ExtractMsgEmailBody(Storage.Message message, ReaderHyperLinks hyperlinks, string contentType, bool withHeaderTable = true, bool includeReactionsInfo = false)
        {
            Logger.WriteToLog("Extracting MSG message body from a stream");

            var body = PreProcessMsgFile(message, out var htmlBody);
            if (withHeaderTable)
            {
                var attachments = message?.Attachments?.OfType<Storage.Attachment>().Select(m => m.FileName).ToList();
                var emailHeader = ExtractMsgEmailHeader(message, htmlBody, hyperlinks, attachments, includeReactionsInfo: includeReactionsInfo);
                body = InjectHeader(body, emailHeader, contentType);
            }

            return body;
        }
        #endregion

        #region ReplaceFirstOccurence
        /// <summary>
        /// Method to replace the first occurrence of the <paramref name="search"/> string with a
        /// <paramref name="replace"/> string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="search"></param>
        /// <param name="replace"></param>
        /// <returns></returns>
        private static string ReplaceFirstOccurrence(string text, string search, string replace)
        {
            var index = text.IndexOf(search, StringComparison.Ordinal);
            if (index < 0)
                return text;

            return text.Substring(0, index) + replace + text.Substring(index + search.Length);
        }
        #endregion

        #region ExtractMsgEmailHeader
        /// <summary>
        /// Returns the header information from the given e-mail <paramref name="message"/> 
        /// (not Appointments, Tasks, Contacts and Sticky notes!!)
        /// </summary>
        /// <param name="message">The message</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        public string ExtractMsgEmailHeader(Storage.Message message, ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Extracting MSG header");

            var htmlBody = false;

            if (string.IsNullOrEmpty(message.BodyHtml))
            {
                // Can still be converted to HTML
                if (message.BodyRtf != null)
                    htmlBody = true;
                // If there is also no text body then we generate a default html body
                else if (string.IsNullOrEmpty(message.BodyText))
                    htmlBody = true;
            }

            var attachmentList = new List<string>();

            foreach (var attachment in message.Attachments)
            {
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    if (!attach.IsInline)
                        attachmentList.Add(attach.FileName);
                }
                // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
                else if (attachment is Storage.Message)
                // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
                {
                    var msg = (Storage.Message)attachment;
                    if (msg.RenderingPosition != -1)
                        attachmentList.Add(msg.FileName);
                }
            }

            return ExtractMsgEmailHeader(message, htmlBody, hyperlinks, attachmentList);
        }

        /// <summary>
        /// Returns the header information from the given e-mail <paramref name="message"/> 
        /// (not Appointments, Tasks, Contacts and Sticky notes!!)
        /// </summary>
        /// <param name="message">The message</param>
        /// <param name="htmlBody">Indicates that the message has an HTML body</param>
        /// <param name="hyperlinks">When set to true then hyperlinks are generated for To, CC and BCC</param>
        /// <param name="attachmentList">A list with attachments</param>
        /// <param name="includeReactionsInfo">When <c>true</c> then reactions information is also included in the output</param>
        /// <returns></returns>
        private string ExtractMsgEmailHeader(Storage.Message message,
                                             bool htmlBody,
                                             ReaderHyperLinks hyperlinks,
                                             List<string> attachmentList = null,
                                             bool includeReactionsInfo = false)
        {
            Logger.WriteToLog("Extracting MSG header");

            var convertToHref = false;

            if (htmlBody)
            {
                switch (hyperlinks)
                {
                    case ReaderHyperLinks.Email:
                        convertToHref = true;
                        break;
                    case ReaderHyperLinks.Both:
                        convertToHref = true;
                        break;
                }
            }

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.EmailSignedBy,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    LanguageConsts.EmailFollowUpFlag,
                    LanguageConsts.EmailFollowUpLabel,
                    LanguageConsts.EmailFollowUpStatusLabel,
                    LanguageConsts.EmailFollowUpCompletedText,
                    LanguageConsts.TaskStartDateLabel,
                    LanguageConsts.TaskDueDateLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel,
                    LanguageConsts.ReactionsSummary,
                    LanguageConsts.OwnerReactionHistory
                    #endregion
                };

                if (message.Type is MessageType.EmailEncryptedAndMaybeSigned or MessageType.EmailClearSigned)
                    languageConsts.Add(LanguageConsts.EmailSignedBy);

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var emailHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel,
                message.GetEmailSender(htmlBody, convertToHref));

            // Sent on
            if (message.SentOn != null)
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                    ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailRecipients(RecipientType.To, htmlBody, convertToHref));

            // CC
            var cc = message.GetEmailRecipients(RecipientType.Cc, htmlBody, convertToHref);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailRecipients(RecipientType.Bcc, htmlBody, convertToHref);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            if (message.Type is MessageType.EmailEncryptedAndMaybeSigned or MessageType.EmailClearSigned)
            {
                var signerInfo = message.SignedBy;
                if (message.SignedOn != null)
                {
                    signerInfo += " " + LanguageConsts.EmailSignedByOn + " " +
                                  ((DateTime)message.SignedOn).ToString(LanguageConsts.DataFormatWithTime);

                    WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSignedBy,
                        signerInfo);
                }
            }

            // Subject
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, message.Subject);

            // Urgent
            if (!string.IsNullOrEmpty(message.ImportanceText))
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, message.ImportanceText);

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Attachments
            if (attachmentList != null && attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                string.Join(", ", attachmentList));

            // ReactionsSummary
            if (includeReactionsInfo)
            {
                var currentReactions = message.GetCurrentReactionStringList();
                if (currentReactions != null && currentReactions.Count != 0)
                    WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.ReactionsSummary, string.Join("\n", currentReactions));

                // OwnerReactionHistory
                var ownerReactions = message.GetOwnerReactionStringList();
                if (ownerReactions != null && ownerReactions.Count != 0)
                    WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.OwnerReactionHistory, string.Join("\n", ownerReactions));

            }

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // Follow up
            if (message.Flag != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpLabel,
                    message.Flag.Request);

                if (message.Task != null)
                {
                    // When complete
                    if (message.Task.Complete != null && (bool)message.Task.Complete)
                    {
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpStatusLabel,
                            LanguageConsts.EmailFollowUpCompletedText);

                        // Task completed date
                        if (message.Task.CompleteTime != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDateCompleted,
                                ((DateTime)message.Task.CompleteTime).ToString(LanguageConsts.DataFormatWithTime));
                    }
                    else
                    {
                        // Task start date
                        if (message.Task.StartDate != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskStartDateLabel,
                                ((DateTime)message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

                        // Task due date
                        if (message.Task.DueDate != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDueDateLabel,
                                ((DateTime)message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));
                    }
                }

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Categories
            var categories = message.Categories;
            if (categories != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    string.Join("; ", categories));

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);
            return emailHeader.ToString();
        }
        #endregion

        #region WriteHeader methods
        /// <summary>
        /// Surrounds the String with HTML tags
        /// </summary>
        /// <param name="footer"></param>
        /// <param name="htmlBody"></param>
        private static void SurroundWithHtml(StringBuilder footer, bool htmlBody)
        {
            if (!htmlBody)
                return;
            footer.Insert(0, "<html><body><br/>");
            footer.AppendLine("</body></html>");

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes the start of the header
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <param ref="header"/> otherwise text will be written</param>
        private void WriteHeaderStart(StringBuilder header, bool htmlBody)
        {
            if (!htmlBody)
                return;

            if (UseCustomHeaderStyle)
            {
                header.AppendLine("<style>" + GetCustomHeaderStyle() + "</style>");
                header.AppendLine("<table class=\"MsgReaderHeader\">");
            }
            else
                header.AppendLine("<table style=\"font-family: Times New Roman; font-size: 12pt;\">");

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes a line into the header
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <paramref name="header"/> otherwise text will be written</param>
        /// <param name="labelPadRightWidth">Used to pad the label size, ignored when <paramref name="htmlBody"/> is true</param>
        /// <param name="label">The label text that needs to be written</param>
        /// <param name="text">The text that needs to be written after the <paramref name="label"/></param>
        private void WriteHeaderLine(StringBuilder header,
            bool htmlBody,
            int labelPadRightWidth,
            string label,
            string text)
        {
            if (htmlBody)
            {
                var lines = text.Split('\n');
                var newText = string.Empty;

                foreach (var line in lines)
                    newText += WebUtility.HtmlEncode(line) + "<br/>";

                string htmlTr;

                if (UseCustomHeaderStyle)
                {
                    htmlTr =
                        "<tr class=\"MsgReaderHeaderRow\">" +
                        "<td class=\"MsgReaderHeaderRowLabel\">";
                }
                else
                {
                    htmlTr =
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"font-weight: bold; white-space:nowrap;\">";
                }

                htmlTr += WebUtility.HtmlEncode(label) + ":</td>" +
                          "<td class=\"MsgReaderHeaderRowText\">" + newText + "</td></tr>";

                header.AppendLine(htmlTr);
            }
            else
            {
                text = text.Replace("\n", "".PadRight(labelPadRightWidth));
                header.AppendLine((label + ":").PadRight(labelPadRightWidth) + text);
            }

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes a line into the header without Html encoding the <paramref name="text"/>
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <paramref name="header"/> otherwise text will be written</param>
        /// <param name="labelPadRightWidth">Used to pad the label size, ignored when <paramref name="htmlBody"/> is true</param>
        /// <param name="label">The label text that needs to be written</param>
        /// <param name="text">The text that needs to be written after the <paramref name="label"/></param>
        private void WriteHeaderLineNoEncoding(StringBuilder header,
                                               bool htmlBody,
                                               int labelPadRightWidth,
                                               string label,
                                               string text)
        {
            if (htmlBody)
            {
                text = text.Replace("\n", "<br/>");

                string htmlTr;

                if (UseCustomHeaderStyle)
                {
                    htmlTr =
                        "<tr class=\"MsgReaderHeaderRow\">" +
                        "<td class=\"MsgReaderHeaderRowLabel\">";
                }
                else
                {
                    htmlTr =
                        "<tr style=\"height: 18px; vertical-align: top; \">" +
                        "<td style=\"font-weight: bold; white-space:nowrap;\">";
                }

                htmlTr += WebUtility.HtmlEncode(label) + ":</td>" +
                          "<td class=\"MsgReaderHeaderRowText\">" + text + "</td>" +
                          "</tr>";

                header.AppendLine(htmlTr);
            }
            else
            {
                text = text.Replace("\n", "".PadRight(labelPadRightWidth));
                header.AppendLine((label + ":").PadRight(labelPadRightWidth) + text);
            }

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes an empty header line
        /// </summary>
        /// <param name="header"></param>
        /// <param name="htmlBody"></param>
        private void WriteHeaderEmptyLine(StringBuilder header, bool htmlBody)
        {
            // Prevent that we write 2 empty lines in a row
            if (_emptyLineWritten)
                return;

            if (!htmlBody)
                header.AppendLine(string.Empty);
            else
            {

                header.AppendLine(
                    UseCustomHeaderStyle
                        ? "<tr class=\"MsgReaderHeaderRow MsgReaderHeaderRowEmpty\">" +
                          "<td class=\"MsgReaderHeaderRowLabel\">&nbsp;</td>" +
                          "<td class=\"MsgReaderHeaderRowText\">&nbsp;</td>" +
                          "</tr>"
                        : "<tr style=\"height: 18px; vertical-align: top; \">" +
                          "<td>&nbsp;</td>" +
                          "<td>&nbsp;</td>" +
                          "</tr>");
            }

            _emptyLineWritten = true;
        }

        /// <summary>
        /// Writes the end of the header
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <param ref="header"/> otherwise text will be written</param>
        private static void WriteHeaderEnd(StringBuilder header, bool htmlBody)
        {
            header.AppendLine(!htmlBody ? string.Empty : "</table><br/>");
        }
        #endregion

        #region WriteMsgEmail
        /// <summary>
        /// Writes the body of the MSG E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <param name="includeReactionsInfo">When <c>true</c> then reactions information is also included in the output</param>
        /// <returns></returns>
        private List<string> WriteMsgEmail(Storage.Message message, string outputFolder, ReaderHyperLinks hyperlinks, bool includeReactionsInfo = false)
        {
            var fileName = "email";

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out var htmlBody,
                out var body,
                out _,
                out var attachmentList,
                out var files);

            var emailHeader = ExtractMsgEmailHeader(message, htmlBody, hyperlinks, attachmentList, includeReactionsInfo);
            body = InjectHeader(body, emailHeader);

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            return files;
        }
        #endregion

        #region WriteEmlStreamEmail
        /// <summary>
        /// Writes the body of the MSG E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of MemoryStream
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        public List<MemoryStream> WriteEmlStreamEmail(Mime.Message message, bool hyperlinks)
        {
            Logger.WriteToLog("Writing EML message body to stream");

            var streams = new List<MemoryStream>();

            PreProcessEmlStream(message,
                hyperlinks,
                out var htmlBody,
                out var body,
                out var attachmentList,
                out var attachStreams);

            if (!htmlBody)
                hyperlinks = false;

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            /*******************************Start Header*******************************/
            Logger.WriteToLog("Start writing EML header information");

            var emailHeader = new StringBuilder();
            var headers = message.Headers;

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            var from = string.Empty;
            if (headers.From != null)
                from = message.GetEmailAddresses(new List<RfcMailAddress> { headers.From }, hyperlinks, htmlBody);

            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel, from);

            // Sent on
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                message.Headers.DateSent.ToLocalTime().ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailAddresses(headers.To, hyperlinks, htmlBody));

            // CC
            var cc = message.GetEmailAddresses(headers.Cc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailAddresses(headers.Bcc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            // Subject
            var subject = message.Headers.Subject ?? string.Empty;
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, subject);

            // Urgent
            var importanceText = string.Empty;
            switch (message.Headers.Importance)
            {
                case MailPriority.Low:
                    importanceText = LanguageConsts.ImportanceLowText;
                    break;

                case MailPriority.Normal:
                    importanceText = LanguageConsts.ImportanceNormalText;
                    break;

                case MailPriority.High:
                    importanceText = LanguageConsts.ImportanceHighText;
                    break;
            }

            if (!string.IsNullOrEmpty(importanceText))
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importanceText);

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Attachments
            if (attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                    string.Join(", ", attachmentList));

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            body = InjectHeader(body, emailHeader.ToString());

            var bodyData = Encoding.UTF8.GetBytes(body);
            streams.Add(StreamHelpers.Manager.GetStream("Reader.cs", bodyData, 0, bodyData.Length));

            Logger.WriteToLog("End writing EML header information");

            /*******************************End Header*********************************/

            streams.AddRange(attachStreams);

            /*******************************Start Footer*******************************/
            Logger.WriteToLog("Start writing EML footer information");
            var emailFooter = new StringBuilder();

            WriteHeaderStart(emailFooter, htmlBody);
            var i = 0;
            
            foreach (var item in headers.UnknownHeaders.AllKeys)
            {
                WriteHeaderLine(emailFooter, htmlBody, maxLength, item, headers.UnknownHeaders[i]);
                i++;
            }

            SurroundWithHtml(emailFooter, htmlBody);
            var footerData = Encoding.UTF8.GetBytes(emailFooter.ToString());
            streams.Add(StreamHelpers.Manager.GetStream("Reader.cs", footerData, 0, footerData.Length));

            Logger.WriteToLog("End writing EML footer information");
            /*******************************End Header*********************************/

            return streams;
        }
        #endregion

        #region WriteEmlEmail
        /// <summary>
        /// Writes the body of the EML E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <returns></returns>
        private List<string> WriteEmlEmail(Mime.Message message, string outputFolder, ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Start writing EML e-mail body and attachments to outputfolder");

            var fileName = "email";

            PreProcessEmlFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out var htmlBody,
                out var body,
                out var attachmentList,
                out var files);

            var convertToHref = false;

            if (htmlBody)
            {
                switch (hyperlinks)
                {
                    case ReaderHyperLinks.Email:
                        convertToHref = true;
                        break;
                    case ReaderHyperLinks.Both:
                        convertToHref = true;
                        break;
                }
            }

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    LanguageConsts.EmailSignedBy,
                    LanguageConsts.EmailSignedByOn
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            Logger.WriteToLog("Start writing EML headers");

            var emailHeader = new StringBuilder();

            var headers = message.Headers;

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            var from = string.Empty;
            if (headers.From != null)
                from = message.GetEmailAddresses(new List<RfcMailAddress> { headers.From }, convertToHref, htmlBody);

            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel, from);

            // Sent on
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                message.Headers.DateSent.ToLocalTime().ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailAddresses(headers.To, convertToHref, htmlBody));

            // CC
            var cc = message.GetEmailAddresses(headers.Cc, convertToHref, htmlBody);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailAddresses(headers.Bcc, convertToHref, htmlBody);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            if (message.SignedBy != null)
            {
                var signerInfo = message.SignedBy;
                if (message.SignedOn != null)
                {
                    signerInfo += " " + LanguageConsts.EmailSignedByOn + " " +
                                  ((DateTime)message.SignedOn).ToString(LanguageConsts.DataFormatWithTime);

                    WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSignedBy,
                        signerInfo);
                }
            }

            // Subject
            var subject = message.Headers.Subject ?? string.Empty;
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, subject);

            // Urgent
            var importanceText = string.Empty;
            switch (message.Headers.Importance)
            {
                case MailPriority.Low:
                    importanceText = LanguageConsts.ImportanceLowText;
                    break;

                case MailPriority.Normal:
                    importanceText = LanguageConsts.ImportanceNormalText;
                    break;

                case MailPriority.High:
                    importanceText = LanguageConsts.ImportanceHighText;
                    break;
            }

            if (!string.IsNullOrEmpty(importanceText))
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importanceText);

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Attachments
            if (attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                    string.Join(", ", attachmentList));

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            Logger.WriteToLog("Stop writing EML headers");

            body = InjectHeader(body, emailHeader.ToString());

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            Logger.WriteToLog("Stop writing EML e-mail body and attachments to outputfolder");

            return files;
        }
        #endregion

        #region WriteMsgAppointment
        /// <summary>
        /// Writes the body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <returns></returns>
        private List<string> WriteMsgAppointment(Storage.Message message, string outputFolder, ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Start writing MSG appointment and attachments to outputfolder");

            var fileName = "appointment";

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out var htmlBody,
                out var body,
                out _,
                out var attachmentList,
                out var files);

            var convertToHref = false;

            if (htmlBody)
            {
                switch (hyperlinks)
                {
                    case ReaderHyperLinks.Email:
                        convertToHref = true;
                        break;
                    case ReaderHyperLinks.Both:
                        convertToHref = true;
                        break;
                }
            }

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.AppointmentSubjectLabel,
                    LanguageConsts.AppointmentLocationLabel,
                    LanguageConsts.AppointmentStartDateLabel,
                    LanguageConsts.AppointmentEndDateLabel,
                    LanguageConsts.AppointmentRecurrenceTypeLabel,
                    LanguageConsts.AppointmentClientIntentLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentRecurrencePaternLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentMandatoryParticipantsLabel,
                    LanguageConsts.AppointmentOptionalParticipantsLabel,
                    LanguageConsts.AppointmentCategoriesLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var appointmentHeader = new StringBuilder();

            Logger.WriteToLog("Start writing MSG header");

            // Start of table
            WriteHeaderStart(appointmentHeader, htmlBody);

            // Subject
            if (!string.IsNullOrEmpty(message.Subject))
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentSubjectLabel,
                message.Subject);

            // Location
            if (!string.IsNullOrEmpty(message.Appointment?.Location))
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentLocationLabel,
                message.Appointment.Location);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Start
            if (message.Appointment?.Start != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentStartDateLabel,
                    ((DateTime)message.Appointment.Start).ToString(LanguageConsts.DataFormatWithTime));

            // End
            if (message.Appointment?.End != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentEndDateLabel,
                    ((DateTime)message.Appointment.End).ToString(LanguageConsts.DataFormatWithTime));

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Recurrence type
            if (!string.IsNullOrEmpty(message.Appointment?.RecurrenceTypeText))
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrenceTypeLabel,
                    message.Appointment.RecurrenceTypeText);

            // Recurrence pattern
            if (!string.IsNullOrEmpty(message.Appointment?.RecurrencePattern))
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrencePaternLabel,
                    message.Appointment.RecurrencePattern);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Status
            if (!string.IsNullOrEmpty(message.Appointment?.ClientIntentText))
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentClientIntentLabel,
                    message.Appointment.ClientIntentText);

            // Appointment organizer (FROM)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentOrganizerLabel,
                message.GetEmailSender(htmlBody, convertToHref));

            // Mandatory participants (TO)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                LanguageConsts.AppointmentMandatoryParticipantsLabel,
                message.GetEmailRecipients(RecipientType.To, htmlBody, convertToHref));

            // Optional participants (CC)
            var cc = message.GetEmailRecipients(RecipientType.Cc, htmlBody, convertToHref);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentOptionalParticipantsLabel, cc);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Categories
            var categories = message.Categories;
            if (categories != null)
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    string.Join("; ", categories));

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Urgent
            var importance = message.ImportanceText;
            if (!string.IsNullOrEmpty(importance))
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importance);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Attachments
            if (attachmentList.Count != 0)
            {
                WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentAttachmentsLabel,
                    string.Join(", ", attachmentList));

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // End of table + empty line
            WriteHeaderEnd(appointmentHeader, htmlBody);

            body = InjectHeader(body, appointmentHeader.ToString());
            Logger.WriteToLog("Stop writing MSG header");

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            Logger.WriteToLog("Stop writing MSG appointment and attachments to outputfolder");
            return files;
        }
        #endregion

        #region WriteMsgTask
        /// <summary>
        /// Writes the task body of the MSG Task to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <returns></returns>
        private List<string> WriteMsgTask(Storage.Message message, string outputFolder, ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Start writing MSG task and attachments to outputfolder");

            var fileName = "task";

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out var htmlBody,
                out var body,
                out _,
                out var attachmentList,
                out var files);

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.TaskSubjectLabel,
                    LanguageConsts.TaskStartDateLabel,
                    LanguageConsts.TaskDueDateLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskStatusLabel,
                    LanguageConsts.TaskPercentageCompleteLabel,
                    LanguageConsts.TaskEstimatedEffortLabel,
                    LanguageConsts.TaskActualEffortLabel,
                    LanguageConsts.TaskOwnerLabel,
                    LanguageConsts.TaskContactsLabel,
                    LanguageConsts.EmailCategoriesLabel,
                    LanguageConsts.TaskCompanyLabel,
                    LanguageConsts.TaskBillingInformationLabel,
                    LanguageConsts.TaskMileageLabel
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            Logger.WriteToLog("Start writing MSG header");

            var taskHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(taskHeader, htmlBody);

            // Subject
            WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskSubjectLabel, message.Subject);

            // Task start date
            if (message.Task?.StartDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskStartDateLabel,
                    ((DateTime)message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

            // Task due date
            if (message.Task?.DueDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskDueDateLabel,
                    ((DateTime)message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));

            // Urgent
            var importance = message.ImportanceText;
            if (!string.IsNullOrEmpty(importance))
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importance);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // Status
            if (!string.IsNullOrEmpty(message.Task?.StatusText))
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskStatusLabel, message.Task.StatusText);

            // Percentage complete
            if (message.Task?.PercentageComplete != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskPercentageCompleteLabel,
                    (message.Task.PercentageComplete * 100) + "%");

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // Estimated effort
            if (!string.IsNullOrEmpty(message.Task?.EstimatedEffortText))
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskEstimatedEffortLabel,
                    message.Task.EstimatedEffortText);

                // Actual effort
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskActualEffortLabel,
                    message.Task.ActualEffortText);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Owner
            if (!string.IsNullOrEmpty(message.Task?.Owner))
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskOwnerLabel, message.Task.Owner);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Contacts
            if (message.Task?.Contacts != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskContactsLabel,
                    string.Join("; ", message.Task.Contacts.ToArray()));

            // Categories
            var categories = message.Categories;
            if (categories != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

            // Companies
            if (message.Task?.Companies != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskCompanyLabel,
                    string.Join("; ", message.Task.Companies.ToArray()));

            // Billing information
            if (!string.IsNullOrEmpty(message.Task?.BillingInformation))
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskBillingInformationLabel,
                    message.Task.BillingInformation);

            // Mileage
            if (!string.IsNullOrEmpty(message.Task?.Mileage))
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskMileageLabel, message.Task.Mileage);

            // Attachments
            if (attachmentList.Count != 0)
            {
                WriteHeaderLineNoEncoding(taskHeader, htmlBody, maxLength, LanguageConsts.AppointmentAttachmentsLabel,
                    string.Join(", ", attachmentList));

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // End of table
            WriteHeaderEnd(taskHeader, htmlBody);

            body = InjectHeader(body, taskHeader.ToString());
            Logger.WriteToLog("Stop writing MSG header");

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            Logger.WriteToLog("Stop writing MSG task and attachments to outputfolder");
            return files;
        }
        #endregion

        #region WriteMsgContact
        /// <summary>
        /// Writes the body of the MSG Contact to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <returns></returns>
        private List<string> WriteMsgContact(Storage.Message message, string outputFolder, ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Start writing MSG contact and attachments to outputfolder");

            var fileName = "contact";

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out var htmlBody,
                out var body,
                out var contactPhotoFileName,
                out _,
                out var files);

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                #region Language consts
                var languageConsts = new List<string>
                {
                    LanguageConsts.DisplayNameLabel,
                    LanguageConsts.SurNameLabel,
                    LanguageConsts.GivenNameLabel,
                    LanguageConsts.FunctionLabel,
                    LanguageConsts.DepartmentLabel,
                    LanguageConsts.CompanyLabel,
                    LanguageConsts.WorkAddressLabel,
                    LanguageConsts.BusinessTelephoneNumberLabel,
                    LanguageConsts.BusinessTelephoneNumber2Label,
                    LanguageConsts.BusinessFaxNumberLabel,
                    LanguageConsts.HomeAddressLabel,
                    LanguageConsts.HomeTelephoneNumberLabel,
                    LanguageConsts.HomeTelephoneNumber2Label,
                    LanguageConsts.HomeFaxNumberLabel,
                    LanguageConsts.OtherAddressLabel,
                    LanguageConsts.OtherFaxLabel,
                    LanguageConsts.PrimaryTelephoneNumberLabel,
                    LanguageConsts.PrimaryFaxNumberLabel,
                    LanguageConsts.AssistantTelephoneNumberLabel,
                    LanguageConsts.InstantMessagingAddressLabel,
                    LanguageConsts.CompanyMainTelephoneNumberLabel,
                    LanguageConsts.CellularTelephoneNumberLabel,
                    LanguageConsts.CarTelephoneNumberLabel,
                    LanguageConsts.RadioTelephoneNumberLabel,
                    LanguageConsts.BeeperTelephoneNumberLabel,
                    LanguageConsts.CallbackTelephoneNumberLabel,
                    LanguageConsts.TextTelephoneLabel,
                    LanguageConsts.ISDNNumberLabel,
                    LanguageConsts.TelexNumberLabel,
                    LanguageConsts.Email1EmailAddressLabel,
                    LanguageConsts.Email1DisplayNameLabel,
                    LanguageConsts.Email2EmailAddressLabel,
                    LanguageConsts.Email2DisplayNameLabel,
                    LanguageConsts.Email3EmailAddressLabel,
                    LanguageConsts.Email3DisplayNameLabel,
                    LanguageConsts.BirthdayLabel,
                    LanguageConsts.WeddingAnniversaryLabel,
                    LanguageConsts.SpouseNameLabel,
                    LanguageConsts.ProfessionLabel,
                    LanguageConsts.HtmlLabel
                };
                #endregion

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            Logger.WriteToLog("Start writing MSG header");
            var contactHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(contactHeader, htmlBody);

            if (htmlBody && !string.IsNullOrEmpty(contactPhotoFileName))
            {
                contactHeader.Append(
                    UseCustomHeaderStyle
                        ? "<div class=\"MsgReaderContactPhoto\">" +
                          " <img alt=\"\" src=\"" + contactPhotoFileName + "\">" +
                          "</div>"
                        : "<div style=\"height: 250px; position: absolute; top: 20px; right: 20px;\">" +
                          " <img alt=\"\" src=\"" + contactPhotoFileName + "\" height=\"100%\">" +
                          "</div>"
                );
            }

            // Full name
            if (!string.IsNullOrEmpty(message.Contact.DisplayName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.DisplayNameLabel,
                    message.Contact.DisplayName);

            // Last name
            if (!string.IsNullOrEmpty(message.Contact.SurName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.SurNameLabel, message.Contact.SurName);

            // First name
            if (!string.IsNullOrEmpty(message.Contact.GivenName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.GivenNameLabel, message.Contact.GivenName);

            // Job title
            if (!string.IsNullOrEmpty(message.Contact.Function))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.FunctionLabel, message.Contact.Function);

            // Department
            if (!string.IsNullOrEmpty(message.Contact.Department))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.DepartmentLabel,
                    message.Contact.Department);

            // Company
            if (!string.IsNullOrEmpty(message.Contact.Company))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CompanyLabel, message.Contact.Company);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Business address
            if (!string.IsNullOrEmpty(message.Contact.WorkAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.WorkAddressLabel,
                    message.Contact.WorkAddress);

            // Home address
            if (!string.IsNullOrEmpty(message.Contact.HomeAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeAddressLabel,
                    message.Contact.HomeAddress);

            // Other address
            if (!string.IsNullOrEmpty(message.Contact.OtherAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.OtherAddressLabel,
                    message.Contact.OtherAddress);

            // Instant messaging
            if (!string.IsNullOrEmpty(message.Contact.InstantMessagingAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.InstantMessagingAddressLabel,
                    message.Contact.InstantMessagingAddress);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Business telephone number
            if (!string.IsNullOrEmpty(message.Contact.BusinessTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessTelephoneNumberLabel,
                    message.Contact.BusinessTelephoneNumber);

            // Business telephone number 2
            if (!string.IsNullOrEmpty(message.Contact.BusinessTelephoneNumber2))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessTelephoneNumber2Label,
                    message.Contact.BusinessTelephoneNumber2);

            // Assistant's telephone number
            if (!string.IsNullOrEmpty(message.Contact.AssistantTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.AssistantTelephoneNumberLabel,
                    message.Contact.AssistantTelephoneNumber);

            // Company main phone
            if (!string.IsNullOrEmpty(message.Contact.CompanyMainTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CompanyMainTelephoneNumberLabel,
                    message.Contact.CompanyMainTelephoneNumber);

            // Home telephone number
            if (!string.IsNullOrEmpty(message.Contact.HomeTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeTelephoneNumberLabel,
                    message.Contact.HomeTelephoneNumber);

            // Home telephone number 2
            if (!string.IsNullOrEmpty(message.Contact.HomeTelephoneNumber2))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeTelephoneNumber2Label,
                    message.Contact.HomeTelephoneNumber2);

            // Mobile phone
            if (!string.IsNullOrEmpty(message.Contact.CellularTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CellularTelephoneNumberLabel,
                    message.Contact.CellularTelephoneNumber);

            // Car phone
            if (!string.IsNullOrEmpty(message.Contact.CarTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CarTelephoneNumberLabel,
                    message.Contact.CarTelephoneNumber);

            // Radio
            if (!string.IsNullOrEmpty(message.Contact.RadioTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.RadioTelephoneNumberLabel,
                    message.Contact.RadioTelephoneNumber);

            // Beeper
            if (!string.IsNullOrEmpty(message.Contact.BeeperTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BeeperTelephoneNumberLabel,
                    message.Contact.BeeperTelephoneNumber);

            // Callback
            if (!string.IsNullOrEmpty(message.Contact.CallbackTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CallbackTelephoneNumberLabel,
                    message.Contact.CallbackTelephoneNumber);

            // Other
            if (!string.IsNullOrEmpty(message.Contact.OtherTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.OtherTelephoneNumberLabel,
                    message.Contact.OtherTelephoneNumber);

            // Primary telephone number
            if (!string.IsNullOrEmpty(message.Contact.PrimaryTelephoneNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.PrimaryTelephoneNumberLabel,
                    message.Contact.PrimaryTelephoneNumber);

            // Telex
            if (!string.IsNullOrEmpty(message.Contact.TelexNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.TelexNumberLabel,
                    message.Contact.TelexNumber);

            // TTY/TDD phone
            if (!string.IsNullOrEmpty(message.Contact.TextTelephone))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.TextTelephoneLabel,
                    message.Contact.TextTelephone);

            // ISDN
            if (!string.IsNullOrEmpty(message.Contact.ISDNNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.ISDNNumberLabel,
                    message.Contact.ISDNNumber);

            // Other fax (primary fax, weird that they call it like this in Outlook)
            if (!string.IsNullOrEmpty(message.Contact.PrimaryFaxNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.PrimaryFaxNumberLabel,
                    message.Contact.OtherTelephoneNumber);

            // Business fax
            if (!string.IsNullOrEmpty(message.Contact.BusinessFaxNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessFaxNumberLabel,
                    message.Contact.BusinessFaxNumber);

            // Home fax
            if (!string.IsNullOrEmpty(message.Contact.HomeFaxNumber))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeFaxNumberLabel,
                    message.Contact.HomeFaxNumber);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // E-mail
            if (!string.IsNullOrEmpty(message.Contact.Email1EmailAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email1EmailAddressLabel,
                    message.Contact.Email1EmailAddress);

            // E-mail display as
            if (!string.IsNullOrEmpty(message.Contact.Email1DisplayName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email1DisplayNameLabel,
                    message.Contact.Email1DisplayName);

            // E-mail 2
            if (!string.IsNullOrEmpty(message.Contact.Email2EmailAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email2EmailAddressLabel,
                    message.Contact.Email2EmailAddress);

            // E-mail display as 2
            if (!string.IsNullOrEmpty(message.Contact.Email2DisplayName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email2DisplayNameLabel,
                    message.Contact.Email2DisplayName);

            // E-mail 3
            if (!string.IsNullOrEmpty(message.Contact.Email3EmailAddress))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email3EmailAddressLabel,
                    message.Contact.Email3EmailAddress);

            // E-mail display as 3
            if (!string.IsNullOrEmpty(message.Contact.Email3DisplayName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email3DisplayNameLabel,
                    message.Contact.Email3DisplayName);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Birthday
            if (message.Contact.Birthday != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BirthdayLabel,
                    ((DateTime)message.Contact.Birthday).ToString(LanguageConsts.DataFormat));

            // Anniversary
            if (message.Contact.WeddingAnniversary != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.WeddingAnniversaryLabel,
                    ((DateTime)message.Contact.WeddingAnniversary).ToString(LanguageConsts.DataFormat));

            // Spouse/Partner
            if (!string.IsNullOrEmpty(message.Contact.SpouseName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.SpouseNameLabel,
                    message.Contact.SpouseName);

            // Profession
            if (!string.IsNullOrEmpty(message.Contact.Profession))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.ProfessionLabel,
                    message.Contact.Profession);

            // Assistant
            if (!string.IsNullOrEmpty(message.Contact.AssistantName))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.AssistantTelephoneNumberLabel,
                    message.Contact.AssistantName);

            // Web page
            if (!string.IsNullOrEmpty(message.Contact.Html))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HtmlLabel, message.Contact.Html);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Categories
            var categories = message.Categories;
            if (categories != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    string.Join("; ", categories));

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            WriteHeaderEnd(contactHeader, htmlBody);

            Logger.WriteToLog("Stop writing MSG header");

            body = InjectHeader(body, contactHeader.ToString());

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            Logger.WriteToLog("Stop writing MSG contact and attachments to outputfolder");
            return files;
        }
        #endregion

        #region WriteMsgStickyNote
        /// <summary>
        /// Writes the body of the MSG StickyNote to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <returns></returns>
        private List<string> WriteMsgStickyNote(Storage.Message message, string outputFolder)
        {
            Logger.WriteToLog("Stop writing MSG sticky note to outputfolder");

            var files = new List<string>();
            string stickyNoteFile;
            Logger.WriteToLog("Start writing MSG header");

            var stickyNoteHeader = new StringBuilder();

            // Sticky notes only have RTF or Text bodies
            var body = message.BodyRtf;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                body = RtfToHtmlConverter.ConvertRtfToHtml(body);
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     // ReSharper disable once StringLiteralTypo
                                     : "stickynote") + ".htm";

                WriteHeaderStart(stickyNoteHeader, true);

                if (message.SentOn != null)
                    WriteHeaderLine(stickyNoteHeader, true, 0, LanguageConsts.StickyNoteDateLabel,
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

                // Empty line
                WriteHeaderEmptyLine(stickyNoteHeader, true);

                // End of table + empty line
                WriteHeaderEnd(stickyNoteHeader, true);

                body = InjectHeader(body, stickyNoteHeader.ToString());
            }
            else
            {
                body = message.BodyText ?? string.Empty;

                // Sent on
                if (message.SentOn != null)
                    WriteHeaderLine(stickyNoteHeader, false, LanguageConsts.StickyNoteDateLabel.Length,
                        LanguageConsts.StickyNoteDateLabel,
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

                body = stickyNoteHeader + body;
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     // ReSharper disable once StringLiteralTypo
                                     : "stickynote") + ".txt";
            }

            Logger.WriteToLog("Stop writing MSG header");

            // Write the body to a file
            stickyNoteFile = FileManager.FileExistsMakeNew(stickyNoteFile);
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            files.Add(stickyNoteFile);
            Logger.WriteToLog("Stop writing MSG sticky note to outputfolder");
            return files;
        }
        #endregion

        #region WriteMsgJournal
        /// <summary>
        /// Writes the body of the MSG Journal to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When set to true then hyperlinks are generated for To, CC and BCC</param>
        /// <returns></returns>
        private List<string> WriteMsgJournal(
            Storage.Message message,
            string outputFolder,
            ReaderHyperLinks hyperlinks)
        {
            Logger.WriteToLog("Stop writing MSG journal note to outputfolder");

            var convertToHref = false;

            switch (hyperlinks)
            {
                case ReaderHyperLinks.Email:
                    convertToHref = true;
                    break;
                case ReaderHyperLinks.Both:
                    convertToHref = true;
                    break;
            }

            var files = new List<string>();
            string stickyNoteFile;
            Logger.WriteToLog("Start writing MSG header");

            var journalHeader = new StringBuilder();

            // Sticky notes only have RTF or Text bodies
            var body = message.BodyRtf;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                body = RtfToHtmlConverter.ConvertRtfToHtml(body);
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     // ReSharper disable once StringLiteralTypo
                                     : "journal") + ".htm";

                WriteHeaderStart(journalHeader, true);

                // From
                WriteHeaderLineNoEncoding(journalHeader, true, 0, LanguageConsts.EmailFromLabel, message.GetEmailSender(true, convertToHref));

                if (message.SentOn != null)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.StickyNoteDateLabel,
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

                // Subject
                WriteHeaderLine(journalHeader, true, 0, LanguageConsts.EmailSubjectLabel, message.Subject);

                // Empty line
                WriteHeaderEmptyLine(journalHeader, true);

                WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogType, message.Log.Type);
                WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogTypeDescription, message.Log.TypeDescription);

                // Empty line
                WriteHeaderEmptyLine(journalHeader, true);

                if (message.Log.Start.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogStart,
                        ((DateTime)message.Log.Start).ToString(LanguageConsts.DataFormatWithTime));

                if (message.Log.End.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogEnd,
                        ((DateTime)message.Log.End).ToString(LanguageConsts.DataFormatWithTime));

                if (message.Log.Duration.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogDuration, message.Log.Duration.ToString());

                // Empty line
                WriteHeaderEmptyLine(journalHeader, true);

                if (message.Log.DocumentPrinted.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogDocumentPrinted, message.Log.DocumentPrinted.ToString());
                if (message.Log.DocumentSaved.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogDocumentSaved, message.Log.DocumentSaved.ToString());
                if (message.Log.DocumentRouted.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogDocumentRouted, message.Log.DocumentRouted.ToString());
                if (message.Log.DocumentPosted.HasValue)
                    WriteHeaderLine(journalHeader, true, 0, LanguageConsts.LogDocumentPosted, message.Log.DocumentPosted.ToString());

                // End of table + empty line
                WriteHeaderEnd(journalHeader, true);

                body = InjectHeader(body, journalHeader.ToString());
            }
            else
            {
                body = message.BodyText ?? string.Empty;

                // Sent on
                if (message.SentOn != null)
                    WriteHeaderLine(journalHeader, false, LanguageConsts.StickyNoteDateLabel.Length,
                        LanguageConsts.StickyNoteDateLabel,
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

                body = journalHeader + body;
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     // ReSharper disable once StringLiteralTypo
                                     : "journal") + ".txt";
            }

            Logger.WriteToLog("Stop writing MSG header");

            // Write the body to a file
            stickyNoteFile = FileManager.FileExistsMakeNew(stickyNoteFile);
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            files.Add(stickyNoteFile);
            Logger.WriteToLog("Stop writing MSG journal to outputfolder");
            return files;
        }
        #endregion

        #region PreProcessMsgFile
        /// <summary>
        /// This method reads the body of a message object and returns it as a html body
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="htmlBody">Returns <c>true</c> when a html body is returned, <c>false</c>
        /// when the body is text based</param>
        /// <returns>True when the e-Mail has an HTML body</returns>
        private static string PreProcessMsgFile(Storage.Message message, out bool htmlBody)
        {
            Logger.WriteToLog("Start pre processing MSG file");

            // ReSharper disable once StringLiteralTypo
            const string rtfInlineObject = "[*[RTFINLINEOBJECT]*]";

            htmlBody = true;
            var body = message.BodyHtml;
            if (string.IsNullOrEmpty(body))
            {
                htmlBody = false;
                body = message.BodyRtf;

                // If the body is not null then we convert it to HTML
                if (body != null)
                {
                    // The RtfToHtmlConverter doesn't support the RTF \objattph tag. So we need to 
                    // replace the tag with some text that does survive the conversion. Later on we 
                    // will replace these tags with the correct inline image tags
                    body = body.Replace("\\objattph", rtfInlineObject);
                    body = RtfToHtmlConverter.ConvertRtfToHtml(body);
                    htmlBody = true;
                }
                else
                {
                    body = message.BodyText;

                    // When there is not a body at all we just make an empty html document
                    if (body == null)
                    {
                        htmlBody = true;
                        body = "<html><head></head><body></body></html>";
                    }
                }
            }

            Logger.WriteToLog("Stop pre processing MSG file");
            return body;
        }

        /// <summary>
        /// This function preprocesses the Outlook MSG <see cref="Storage.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Storage.Attachment"/> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and 
        /// attachments (when there is a html body)</param>
        /// <param name="outputFolder">The output folder where all extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Storage.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="contactPhotoFileName">Returns the filename of the contact photo. This field will only
        /// return a value when the <see cref="Storage.Message"/> object is a <see cref="MessageType.Contact"/> 
        /// type and the <see cref="Storage.Message.Attachments"/> contains an object that has the 
        /// <param ref="Storage.Message.Attachment.IsContactPhoto"/> set to true, otherwise this field will always be null</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after preprocessing the <see cref="Storage.Message"/> object</param>
        private void PreProcessMsgFile(Storage.Message message,
            ReaderHyperLinks hyperlinks,
            string outputFolder,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out string contactPhotoFileName,
            out List<string> attachments,
            out List<string> files)
        {
            Logger.WriteToLog("Start pre processing MSG file");
            const string rtfInlineObject = "[*[RTFINLINEOBJECT]*]";

            htmlBody = true;
            attachments = new List<string>();
            files = new List<string>();
            contactPhotoFileName = null;
            body = message.BodyHtml;

            if (string.IsNullOrEmpty(body))
            {
                htmlBody = false;
                Logger.WriteToLog("Getting RTF body");

                body = message.BodyRtf;
                // If the body is not null then we convert it to HTML
                if (body != null)
                {
                    // The RtfToHtmlConverter doesn't support the RTF \objattph tag. So we need to 
                    // replace the tag with some text that does survive the conversion. Later on we 
                    // will replace these tags with the correct inline image tags
                    body = body.Replace("\\objattph", rtfInlineObject);
                    Logger.WriteToLog("Start converting RTF body to HTML");
                    body = RtfToHtmlConverter.ConvertRtfToHtml(body);
                    body = body.Replace("\r\n", "<br>");
                    Logger.WriteToLog("End converting RTF body to HTML");
                    htmlBody = true;
                }
                else
                {
                    Logger.WriteToLog("Getting TEXT body");
                    body = message.BodyText;

                    // When there is not a body at all we just make an empty html document
                    if (body == null)
                    {
                        Logger.WriteToLog("No body found, making an empty HTML body");
                        htmlBody = true;
                        body = "<html><head></head><body></body></html>";
                    }
                }

                Logger.WriteToLog("Stop getting body");
            }

            var subject = FileManager.RemoveInvalidFileNameChars(message.Subject);

            fileName = outputFolder +
                       (!string.IsNullOrEmpty(subject)
                           ? subject
                           : fileName) + (htmlBody ? ".htm" : ".txt");

            fileName = FileManager.FileExistsMakeNew(fileName);
            Logger.WriteToLog($"Body written to '{fileName}'");
            files.Add(fileName);

            Logger.WriteToLog("Start processing attachments");

            var inlineAttachments = new List<InlineAttachment>();

            if (message.Attachments.Count == 0)
            {
                Logger.WriteToLog("Message does not contain any attachments");
            }
            else
            {
                foreach (var attachment in message.Attachments)
                {
                    FileInfo fileInfo = null;
                    var attachmentFileName = string.Empty;
                    var renderingPosition = -1;
                    var isInline = false;

                    // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                    if (attachment is Storage.Attachment)
                    {
                        Logger.WriteToLog("Attachment is of the type Storage.Attachment");

                        var attach = (Storage.Attachment)attachment;
                        if (attach.Data == null) continue;
                        attachmentFileName = attach.FileName;
                        renderingPosition = attach.RenderingPosition;
                        fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attachmentFileName));
                        File.WriteAllBytes(fileInfo.FullName, attach.Data);
                        isInline = attach.IsInline;

                        if (attach.IsContactPhoto && htmlBody)
                        {
                            contactPhotoFileName = fileInfo.FullName;
                            continue;
                        }

                        // When we find an inline attachment we have to replace the CID tag inside the html body
                        // with the name of the inline attachment. But before we do this we check if the CID exists.
                        // When the CID does not exist we treat the inline attachment as a normal attachment
                        if (htmlBody && !string.IsNullOrEmpty(attach.ContentId))
                        {
                            if (body.Contains($"cid:{attach.ContentId}"))
                            {
                                Logger.WriteToLog("Attachment is inline, found by content id");
                                body = body.Replace($"cid:{attach.ContentId}", fileInfo.Name);
                            }
                            else if (body.Contains($"cid:{attach.FileName}"))
                            {
                                Logger.WriteToLog("Attachment is inline, found by filename");
                                body = body.Replace("cid:" + attach.FileName, fileInfo.Name);
                            }
                            else
                            {
                                isInline = false;
                                Logger.WriteToLog($"Attachment was marked as inline but the body did not contain the content id 'cid:{attach.ContentId}' or 'cid:{attach.FileName}' so mark it as a normal attachment");
                            }
                        }
                        else
                        {
                            // If we didn't find the cid tag we treat the inline attachment as a normal one 
                            isInline = false;
                            Logger.WriteToLog(
                                $"Attachment was marked as inline but the body did not contain the content id '{attach.ContentId}' so mark it as a normal attachment");
                        }
                    }
                    // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
                    else if (attachment is Storage.Message)
                    // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
                    {
                        Logger.WriteToLog("Attachment is of the type Storage.Message");
                        var msg = (Storage.Message)attachment;
                        attachmentFileName = msg.FileName;
                        renderingPosition = msg.RenderingPosition;
                        fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attachmentFileName));
                        msg.Save(fileInfo.FullName);
                    }

                    if (fileInfo == null) continue;

                    if (!isInline)
                        files.Add(fileInfo.FullName);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (htmlBody && renderingPosition != -1 && body.Contains(rtfInlineObject))
                    {
                        if (!isInline)
                        {
                            var iconFileName = $"{outputFolder}{Guid.NewGuid()}.png";

#if (NETFRAMEWORK)

                            using var icon = Icon.ExtractAssociatedIcon(fileInfo.FullName);
                            using var iconStream = StreamHelpers.Manager.GetStream();
                            icon?.Save(iconStream);
                            using var image = Image.FromStream(iconStream);
                            using var fileStream = File.OpenWrite(iconFileName);
                            image.Save(fileStream, ImageFormat.Png);

#else

                            const string base64Image = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFK0lEQVR42sWXW0wjVRjHv1OgsNrtiCKJBnahF7ILBNbb4m25w4Nvvmw0Puy+mRBKAmyNWX3zwURuiYlxH3zw2Zc1Go3ReNsrhWJclLLLtjMjGKAr7tKyEOhMz/E7M1Nati096IMTPr4zX5ue///3zZmZQ+B/Pkiu4vj4eElxcXE1BhH5EUopaJq2Ojg4uPmfBYyMjNgrKyunm5ubmwghjPHvMJ6M7/IR4Sl1zjMW2fLy8s6NG7Of2O0l5/x+/86/FjA6Onqst7d3vqGhwXBmzm0ee8fGf7ASqKoCiqrC7YWFL3RdPz08PCwkIksA4j/e09MzV1dXR6K3PmOO0h0koOFMOn5Xx6zhuW6dW3U8X4wCST76OqOMkWtXrwiLyCcg5PV6YXvxU5DKsK1Mw08SZqZWZgkrzHFkpRTok2/Dw47DgO0AURE5BXR398x567xkR73ApEMb2GONEZbYN8srh4jd/T6rqKgg99Zj/JoQIpFHQLdBYEf9EAnE9jjNlyOrDrB7xqGqqgq2t7cBRQiRyCmgq7t7DgWQhPwBk8ruiRGIOond+xGrrq42VgWK2EMikdBO+/3ndsQEdHECHkjI7yGBu2mnND+B2FYxBO76obG5Zfe3OIn4xn2IRqMwMxP8+J3z5/sEBXTNeTxeosnvMqn0LyECPK9v2tmdmMNaLebq2NIfJ+zIOFtaWkxEIpHKoaGheEEBnZ1dBgEt7Aep9I7QNZA7U4hpLqANkxAJR+D65PX6AZ9vXkBAp0FAD/uYZF8VJpDOScz8DkkYCiC08RqTIzIJBCbr+/v7hQSE3B4PJBfeRALLe3sP+10LunGLxInNmzTmmI4EGq+CLEcgEAiICejo6EACHpJcOIsE/hR0bv4Bf1ZQno0HCYvrtSTZdMUgMDUlIGBsjBPoMAjQW6+BVLJUoNfUcGu6TjvnwZ8XMb0WaNNlJCBDAAX4CgsYQwL8GnATevNVJpX8kcexbjjkTsFyzgkQy3mqHtNrSPLEzwwFkKnAVL3PJyTAJMDmXwGpWHmg99l9znT+YD2WrAH9xE+gIAHeAl+hVcAFtOM14Ha7CYS6kEDEcqxhppYz0yHStxxbvWfp3qfq8WQN0Z/6gcmKTKanpgQFtCMBtxsg9BISCKOTZN4+Z9Vpuv8GAYoEnv4eFEWGAwhon3MhATLXwqQieddhZo8znefqfapuEHjmO6YoCpmeFhTQ1tZuECC/t4CzSCnY5/3qnID27LcGgeD0tKiANpPAby8gAcV0SNMOTcfZznPV47SGJJ77hqlIIBgUFNDa1hZyudxgm30RpCJ1t59Z65zmqWeuAnYUtJNfIwEVZoQFtCIBl4vYZl9mTpuavuozHO5ZDak6za7H4ShJnPyK4UsrEggWfhhxAadaW5GAC2y/toJkUwv2Obtm1XFFoABIPP8lqEggOCMq4FSrQWDzlzeYg6ykHQHscShS34QniL3lgkFghhMYGNhfQF9f37EzZ87Oe+vqgCaT5nu/YQd2x+ntgdiYEBtEImH4/OLF47jvuJlLQDGGHaOkvLzc4fe/9SPejr1gejF3RAbX9DkYi93cGVn7E8Iyd0z8mtjdOVF2+dKl8MTERMfa2tp9MO7r/D0f+OYCSjHqMbwYj2Ecxj2hw+l0VuO4yIoSyNiKCWawJkryiMfjS/h2zCffwPgb4zZGiFg/XotxBOMRLgCjDOOhjMm5SBsc7OCLdJu7tERsWedcwDrGIoaSagGxJrNZ7UjlVD01PsjBMiZPjWlGNur/ANao8GmOzufeAAAAAElFTkSuQmCC";
                            var bytes = Convert.FromBase64String(base64Image);
                            File.WriteAllBytes(iconFileName, bytes);

#endif

                            inlineAttachments.Add(new InlineAttachment(iconFileName, attachmentFileName, fileInfo.FullName));
                        }
                        else
                        {
                            inlineAttachments.Add(new InlineAttachment(renderingPosition, attachmentFileName));
                        }
                    }
                    else
                        renderingPosition = -1;

                    if (!isInline && renderingPosition == -1)
                    {
                        if (htmlBody)
                        {
                            if (hyperlinks is ReaderHyperLinks.Attachments or ReaderHyperLinks.Both)
                                attachments.Add("<a href=\"" + fileInfo.Name + "\">" +
                                                WebUtility.HtmlEncode(attachmentFileName) + "</a> (" +
                                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
                            else
                                attachments.Add(WebUtility.HtmlEncode(attachmentFileName) + " (" +
                                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
                        }
                        else
                            attachments.Add(attachmentFileName + " (" + FileManager.GetFileSizeString(fileInfo.Length) +
                                            ")");
                    }

                    Logger.WriteToLog(
                        $"Attachment written to '{attachmentFileName}' with size '{FileManager.GetFileSizeString(fileInfo.Length)}'");
                }

                Logger.WriteToLog("Stop processing attachments");
            }

            if (htmlBody)
                foreach (var inlineAttachment in inlineAttachments.OrderBy(m => m.RenderingPosition))
                {
                    if (inlineAttachment.IconFileName != null)
                        body = ReplaceFirstOccurrence(body, rtfInlineObject,
                            (UseCustomHeaderStyle
                                ? "<table class=\"MsgReaderInlineAttachment\"><tr><td>"
                                : "<table style=\"width: 70px; display: inline; text-align: center; font-family: Times New Roman; font-size: 12pt;\"><tr><td>")
                            +
                            (hyperlinks == ReaderHyperLinks.Attachments || hyperlinks == ReaderHyperLinks.Both ? "<a href=\"" + inlineAttachment.FullName + "\">" : string.Empty) + "<img alt=\"\" src=\"" +
                            inlineAttachment.IconFileName + "\">" + (hyperlinks == ReaderHyperLinks.Attachments || hyperlinks == ReaderHyperLinks.Both ? "</a>" : string.Empty) + "</td></tr><tr><td>" +
                            WebUtility.HtmlEncode(inlineAttachment.AttachmentFileName) +
                            "</td></tr></table>");
                    else
                        body = ReplaceFirstOccurrence(body, rtfInlineObject, "<img alt=\"\" src=\"" + inlineAttachment.AttachmentFileName + "\">");
                }

            Logger.WriteToLog("Stop pre processing MSG file");
        }
        #endregion

        #region PreProcessEmlStream
        /// <summary>
        /// This function preprocesses the EML <see cref="Mime.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Mime.MessagePart">attachment</see> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and
        /// attachments (when there is a html body)</param>
        /// <param name="htmlBody">Returns true when the <see cref="Mime.Message"/> object did contain
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="attachStreams">Returns all the attachments as a list of streams</param>
        public void PreProcessEmlStream(Mime.Message message,
            bool hyperlinks,
            out bool htmlBody,
            out string body,
            out List<string> attachments,
            out List<MemoryStream> attachStreams)
        {
            Logger.WriteToLog("Start pre processing EML stream");

            attachments = new List<string>();
            attachStreams = new List<MemoryStream>();

            var bodyMessagePart = message.HtmlBody;

            if (bodyMessagePart != null)
            {
                Logger.WriteToLog("Getting HTML body");
                body = bodyMessagePart.GetBodyAsText();
                htmlBody = true;
            }
            else
            {
                bodyMessagePart = message.TextBody;

                // When there is not a body at all we just make an empty html document
                if (bodyMessagePart != null)
                {
                    Logger.WriteToLog("Getting TEXT body");
                    body = bodyMessagePart.GetBodyAsText();
                    htmlBody = false;
                }
                else
                {
                    Logger.WriteToLog("No body found, making an empty HTML body");
                    htmlBody = true;
                    body = "<html><head></head><body></body></html>";
                }
            }

            Logger.WriteToLog("Stop getting body");

            if (message.Attachments != null)
            {
                Logger.WriteToLog("Start processing attachments");
                foreach (var attachment in message.Attachments)
                {
                    var attachmentFileName = attachment.FileName;

                    //use the stream here and don't worry about needing to close it
                    var attachmentBodyData = Encoding.UTF8.GetBytes(body);
                    attachStreams.Add(StreamHelpers.Manager.GetStream("Reader.cs", attachmentBodyData, 0, attachmentBodyData.Length));

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exist we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attachment.ContentId) && body.Contains(attachment.ContentId))
                    {
                        Logger.WriteToLog("Attachment is inline");
                        body = body.Replace("cid:" + attachment.ContentId, CheckValidAttachment(attachmentFileName));
                    }
                    else
                    {
                        // If we didn't find the cid tag we treat the inline attachment as a normal one

                        if (htmlBody)
                        {
                            Logger.WriteToLog($"Attachment was marked as inline but the body did not contain the content id '{attachment.ContentId}' so mark it as a normal attachment");

                            if (hyperlinks)
                                attachments.Add("<a href=\"" + attachmentFileName + "\">" +
                                                WebUtility.HtmlEncode(CheckValidAttachment(attachmentFileName)) + "</a> (" +
                                                FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                            else
                                attachments.Add(WebUtility.HtmlEncode(CheckValidAttachment(attachmentFileName)) + " (" +
                                                FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                        }
                        else
                            attachments.Add(CheckValidAttachment(attachmentFileName) + " (" +
                                            FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                    }

                    Logger.WriteToLog($"Attachment written to '{attachmentFileName}' with size '{FileManager.GetFileSizeString(attachment.Body.Length)}'");
                }

                Logger.WriteToLog("Start processing attachments");
            }
            else
                Logger.WriteToLog("E-mail does not contain any attachments");


            Logger.WriteToLog("Stop pre processing EML stream");
        }
        #endregion

        #region CheckValidAttachment
        /// <summary>
        /// Check for Valid Attachment
        /// </summary>
        /// <param name="attachmentFileName"></param>
        /// <returns></returns>
        public string CheckValidAttachment(string attachmentFileName)
        {
            var filename = attachmentFileName;
            var attachType = Path.GetExtension(attachmentFileName);
            switch (attachType)
            {
                case ".txt":
                case ".rtf":
                case ".doc":
                case ".docx":
                case ".pdf":
                case ".jpg":
                case ".tif":
                case ".tiff":
                case ".png":
                case ".wmf":
                case ".gif":
                    filename = attachmentFileName;
                    break;
                default:
                    filename = filename + " (This attachment is not a supported attachment type.)";
                    break;
            }
            return filename;
        }
        #endregion

        #region PreProcessEmlFile
        /// <summary>
        /// This function preprocesses the EML <see cref="Mime.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Mime.MessagePart">attachment</see> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks"><see cref="ReaderHyperLinks"/></param>
        /// <param name="outputFolder">The output folder where all extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Mime.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after preprocessing the <see cref="Mime.Message"/> object</param>
        private static void PreProcessEmlFile(Mime.Message message,
            ReaderHyperLinks hyperlinks,
            string outputFolder,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out List<string> attachments,
            out List<string> files)
        {
            Logger.WriteToLog("Start pre processing EML file");

            attachments = new List<string>();
            files = new List<string>();

            var bodyMessagePart = message.HtmlBody;

            if (bodyMessagePart != null)
            {
                Logger.WriteToLog("Getting HTML body");
                body = bodyMessagePart.GetBodyAsText();
                htmlBody = true;
            }
            else
            {
                bodyMessagePart = message.TextBody;

                // When there is not a body at all we just make an empty html document
                if (bodyMessagePart != null)
                {
                    Logger.WriteToLog("Getting TEXT body");
                    body = bodyMessagePart.GetBodyAsText();
                    htmlBody = false;
                }
                else
                {
                    Logger.WriteToLog("No body found, making an empty HTML body");
                    htmlBody = true;
                    body = "<html><head></head><body></body></html>";
                }
            }

            var subject = string.Empty;

            if (message.Headers.Subject != null)
                subject = FileManager.RemoveInvalidFileNameChars(message.Headers.Subject);

            fileName = outputFolder +
                       (!string.IsNullOrEmpty(subject)
                           ? subject
                           : fileName) + (htmlBody ? ".htm" : ".txt");

            fileName = FileManager.FileExistsMakeNew(fileName);

            Logger.WriteToLog($"Body written to '{fileName}'");

            files.Add(fileName);

            if (message.Attachments != null)
            {
                foreach (var attachment in message.Attachments)
                {
                    var attachmentFileName = attachment.FileName;
                    var fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attachmentFileName));
                    File.WriteAllBytes(fileInfo.FullName, attachment.Body);

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exist we treat the inline attachment as a normal attachment
                    if (htmlBody && attachment.IsInline &&
                        (!string.IsNullOrEmpty(attachment.ContentId) && body.Contains($"cid:{attachment.ContentId}") ||
                         body.Contains($"cid:{attachment.FileName}")))
                    {
                        Logger.WriteToLog("Attachment is inline");

                        body = !string.IsNullOrEmpty(attachment.ContentId)
                            ? body.Replace("cid:" + attachment.ContentId, fileInfo.Name)
                            : body.Replace("cid:" + attachment.FileName, fileInfo.Name);
                    }
                    else
                    {
                        // If we didn't find the cid tag we treat the inline attachment as a normal one 

                        files.Add(fileInfo.FullName);

                        if (htmlBody)
                        {
                            Logger.WriteToLog($"Attachment was marked as inline but the body did not contain the content id '{attachment.ContentId}' so mark it as a normal attachment");

                            if (hyperlinks == ReaderHyperLinks.Attachments || hyperlinks == ReaderHyperLinks.Both)
                                attachments.Add("<a href=\"" + fileInfo.Name + "\">" +
                                                WebUtility.HtmlEncode(attachmentFileName) + "</a> (" +
                                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
                            else
                                attachments.Add(WebUtility.HtmlEncode(attachmentFileName) + " (" +
                                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
                        }
                        else
                            attachments.Add(attachmentFileName + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
                    }

                    Logger.WriteToLog($"Attachment written to '{attachmentFileName}' with size '{FileManager.GetFileSizeString(attachment.Body.Length)}'");
                }
            }
            else
                Logger.WriteToLog("E-mail does not contain any attachments");

            Logger.WriteToLog("Stop pre processing EML stream");
        }
        #endregion

        #region GetErrorMessage
        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        public string GetErrorMessage()
        {
            return _errorMessage;
        }
        #endregion

        #region InjectHeader
        /// <summary>
        /// Inject an Outlook style header into the top of the html
        /// </summary>
        /// <param name="body"></param>
        /// <param name="header"></param>
        /// <param name="contentType">Content type</param>
        /// <returns></returns>
        private string InjectHeader(string body, string header, string contentType = null)
        {
            Logger.WriteToLog("Start injecting header into body");

            var begin = body.IndexOf("<BODY", StringComparison.InvariantCultureIgnoreCase);

            if (begin <= 0) return header + body;
            begin = body.IndexOf(">", begin, StringComparison.InvariantCultureIgnoreCase);

            if (InjectHeaderAsIFrame)
            {
                header = "<style>iframe::-webkit-scrollbar {display: none;}</style>" +
                         "<iframe id=\"headerframe\" " +
                         " style=\"border:none; width:100%; margin-bottom:5px;\" " +
                         " onload='javascript:(function(o){o.style.height=o.contentWindow.document.body.scrollHeight+\"px\";}(this));' " +  // ensure height is correct
                         " srcdoc='" +
                         "<html style=\"overflow: hidden;\">" +
                         "     <body style=\"margin: 0;\">" + header + "</body>" +
                         "</html>" +
                         "'></iframe>";
            }

            body = body.Insert(begin + 1, header);

            if (!string.IsNullOrWhiteSpace(contentType))
            {
                // Inject content-type:
                const string head = "<head";
                var headBegin = body.IndexOf(head, StringComparison.InvariantCultureIgnoreCase) + head.Length;
                headBegin = body.IndexOf(">", headBegin, StringComparison.InvariantCultureIgnoreCase);

                var contentHeader =
                    $"{Environment.NewLine}<meta http-equiv=\"Content-Type\" content=\"{contentType}\">{Environment.NewLine}";

                body = body.Insert(headBegin + 1, contentHeader);
            }

            Logger.WriteToLog("Stop injecting header into body");
            return body;
        }
        #endregion
    }
}