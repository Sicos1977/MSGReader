using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Threading;
using MsgReader.Exceptions;
using MsgReader.Helpers;
using MsgReader.Localization;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
// ReSharper disable FunctionComplexityOverflow

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader
{
    #region Interface IReader
    /// <summary>
    /// Interface to make Reader class COM exposable
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
        string[] ExtractToFolderFromCom(string inputFile, string outputFolder, bool hyperlinks = false, string culture = "");

        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        string GetErrorMessage();
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
        /// This message can be retreived with the GetErrorMessage. This way we keep .NET exceptions inside
        /// when this code is called from a COM language
        /// </summary>
        private string _errorMessage;

        /// <summary>
        /// Used to keep track if we already did write an empty line
        /// </summary>
        private static bool _emptyLineWritten;
        #endregion

        #region SetCulture
        /// <summary>
        /// Sets the culture that needs to be used to localize the output of this class. 
        /// Default the current system culture is set. When there is no localization available the
        /// default will be used. This will be en-US.
        /// </summary>
        /// <param name="name">The name of the cultere eg. nl-NL</param>
        public void SetCulture(string name)
        {
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
                throw new MRFileTypeNotSupported("Expected .msg or .eml extension on the inputfile");

            extension = extension.ToUpperInvariant();

            using (var fileStream = File.OpenRead(inputFile))
            {
                var header = new byte[2];
                fileStream.Read(header, 0, 2);

                switch (extension)
                {
                    case ".MSG":
                        // Sometimes the email containts an MSG extension and actualy it's an EML.
                        // Most of the times this happens when a user saves the email manually and types 
                        // the filename. To prevent these kind of errors we do a double check to make sure 
                        // the file is realy an MSG file
                        if (header[0] == 0xD0 && header[1] == 0xCF)
                            return ".MSG";

                        return ".EML";

                    case ".EML":
                        // We can't do an extra check overhere because an EML file is text based 
                        return extension;

                    default:
                        throw new MRFileTypeNotSupported("Wrong file extension, expected .msg or .eml");
                }
            }
        }
        #endregion
        
        #region ExtractToStream

        /// <summary>
        /// This method reads the <see cref="inputStream"/> and when the stream is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body so it looks
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachents <br/>
        /// And in the end returns everything to the output stream
        /// </summary>
        /// <param name="inputStream">The msg stream</param>
        /// <param name="hyperlinks">When true hyperlinks are generated for the To, CC, BCC and attachments</param>
        public List<MemoryStream> ExtractToStream(MemoryStream inputStream, bool hyperlinks = false)
        {
            List<MemoryStream> streams = new List<MemoryStream>();
            var message = Mime.Message.Load(inputStream);
            return WriteEmlStreamEmail(message, hyperlinks);
        }
        #endregion

        #region ExtractToFolder
        /// <summary>
        /// This method reads the <paramref name="inputFile"/> and when the file is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body so it looks 
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachents <br/>
        /// And in the end writes everything to the given <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to save the extracted msg file</param>
        /// <param name="hyperlinks">When true hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <param name="culture"></param>
        public string[] ExtractToFolderFromCom(string inputFile, 
                                               string outputFolder, 
                                               bool hyperlinks = false, 
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
                return new string[0];
            }
        }

        /// <summary>
        /// This method reads the <paramref name="inputFile"/> and when the file is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body so it looks 
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachents <br/>
        /// And in the end writes everything to the given <paramref name="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to save the extracted msg file</param>
        /// <param name="hyperlinks">When true hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns>String array containing the full path to the message body and its attachments</returns>
        /// <exception cref="MRFileTypeNotSupported">Raised when the Microsoft Outlook message type is not supported</exception>
        /// <exception cref="MRInvalidSignedFile">Raised when the Microsoft Outlook signed message is invalid</exception>
        /// <exception cref="ArgumentNullException">Raised when the <param ref="inputFile"/> or <param ref="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <param ref="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the <param ref="outputFolder"/> does not exists</exception>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")]
        public string[] ExtractToFolder(string inputFile, string outputFolder, bool hyperlinks = false)
        {
            outputFolder = FileManager.CheckForBackSlash(outputFolder);
            
            _errorMessage = string.Empty;

            var extension = CheckFileNameAndOutputFolder(inputFile, outputFolder);

            switch (extension)
            {
                case ".EML":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    {
                        var message = Mime.Message.Load(stream);
                        return WriteEmlEmail(message, outputFolder, hyperlinks).ToArray();
                    }

                case ".MSG":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    using (var message = new Storage.Message(stream))
                    {
                        switch (message.Type)
                        {
                            case Storage.Message.MessageType.Email:
                            case Storage.Message.MessageType.EmailSms:
                            case Storage.Message.MessageType.EmailNonDeliveryReport:
                            case Storage.Message.MessageType.EmailDeliveryReport:
                            case Storage.Message.MessageType.EmailDelayedDeliveryReport:
                            case Storage.Message.MessageType.EmailReadReceipt:
                            case Storage.Message.MessageType.EmailNonReadReceipt:
                            case Storage.Message.MessageType.EmailEncryptedAndMaybeSigned:
                            case Storage.Message.MessageType.EmailEncryptedAndMaybeSignedNonDelivery:
                            case Storage.Message.MessageType.EmailEncryptedAndMaybeSignedDelivery:
                            case Storage.Message.MessageType.EmailClearSignedReadReceipt:
                            case Storage.Message.MessageType.EmailClearSignedNonDelivery:
                            case Storage.Message.MessageType.EmailClearSignedDelivery:
                            case Storage.Message.MessageType.EmailBmaStub:
                            case Storage.Message.MessageType.CiscoUnityVoiceMessage:
                            case Storage.Message.MessageType.EmailClearSigned:
                            case Storage.Message.MessageType.RightFaxAdv:
                            case Storage.Message.MessageType.SkypeForBusinessMissedMessage:
                            case Storage.Message.MessageType.SkypeForBusinessConversation:
                                return WriteMsgEmail(message, outputFolder, hyperlinks).ToArray();

                            //case Storage.Message.MessageType.EmailClearSigned:
                            //    throw new MRFileTypeNotSupported("A clear signed message is not supported");

                            case Storage.Message.MessageType.Appointment:
                            case Storage.Message.MessageType.AppointmentNotification:
                            case Storage.Message.MessageType.AppointmentSchedule:
                            case Storage.Message.MessageType.AppointmentRequest:
                            case Storage.Message.MessageType.AppointmentRequestNonDelivery:
                            case Storage.Message.MessageType.AppointmentResponse:
                            case Storage.Message.MessageType.AppointmentResponsePositive:
                            case Storage.Message.MessageType.AppointmentResponsePositiveNonDelivery:
                            case Storage.Message.MessageType.AppointmentResponseNegative:
                            case Storage.Message.MessageType.AppointmentResponseNegativeNonDelivery:
                            case Storage.Message.MessageType.AppointmentResponseTentative:
                            case Storage.Message.MessageType.AppointmentResponseTentativeNonDelivery:
                                return WriteMsgAppointment(message, outputFolder, hyperlinks).ToArray();

                            case Storage.Message.MessageType.Contact:
                                return WriteMsgContact(message, outputFolder, hyperlinks).ToArray();

                            case Storage.Message.MessageType.Task:
                            case Storage.Message.MessageType.TaskRequestAccept:
                            case Storage.Message.MessageType.TaskRequestDecline:
                            case Storage.Message.MessageType.TaskRequestUpdate:
                                return WriteMsgTask(message, outputFolder, hyperlinks).ToArray();
                                
                            case Storage.Message.MessageType.StickyNote:
                                return WriteMsgStickyNote(message, outputFolder).ToArray();

                            case Storage.Message.MessageType.Unknown:
                                throw new MRFileTypeNotSupported("Unsupported message type");
                        }
                    }

                    break;
            }

            return new string[0];
        }
        #endregion

        #region ExtractMessageBody
        /// <summary>
        /// Extract a mail body in memory without saving data on the hard drive.
        /// </summary>
        /// <param name="stream">The message as a stream</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <param name="contentType">Content type, e.g. text/html; charset=utf-8</param>
        /// <param name="withHeaderTable">
        /// When true, a text/html table with information of To, CC, BCC and attachments will
        /// be generated and inserted at the top of the text/html document
        /// </param>
        /// <returns>Body as string (can be html code, ...)</returns>
        public string ExtractMsgEmailBody(Stream stream, bool hyperlinks, string contentType, bool withHeaderTable = true)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            // Reset stream to be sure we start at the beginning
            stream.Seek(0, SeekOrigin.Begin);

            using(var message = new Storage.Message(stream))
                return ExtractMsgEmailBody(message, hyperlinks, contentType, withHeaderTable);
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
        /// <returns>Body as string (can be html code, ...)</returns>
        public string ExtractMsgEmailBody(Storage.Message message, bool hyperlinks, string contentType, bool withHeaderTable = true)
        {
            bool htmlBody;

            var body = PreProcessMsgFile(message, out htmlBody);
            if (withHeaderTable)
            {
                var emailHeader = ExtractMsgEmailHeader(message, htmlBody, hyperlinks);
                body = InjectHeader(body, emailHeader, contentType);
            }

            return body;
        }
        #endregion

        #region ReplaceFirstOccurence
        /// <summary>
        /// Method to replace the first occurence of the <paramref name="search"/> string with a
        /// <paramref name="replace"/> string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="search"></param>
        /// <param name="replace"></param>
        /// <returns></returns>
        private static string ReplaceFirstOccurence(string text, string search, string replace)
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
        /// <param name="hyperlinks">When set to true then hyperlinks are generated for To, CC and BCC</param>
        public string ExtractMsgEmailHeader(Storage.Message message, bool hyperlinks)
        {
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
        /// <returns></returns>
        private string ExtractMsgEmailHeader(Storage.Message message,
                                             bool htmlBody,
                                             bool hyperlinks,
                                             List<string> attachmentList = null)
        {
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
                    LanguageConsts.EmailFollowUpFlag,
                    LanguageConsts.EmailFollowUpLabel,
                    LanguageConsts.EmailFollowUpStatusLabel,
                    LanguageConsts.EmailFollowUpCompletedText,
                    LanguageConsts.TaskStartDateLabel,
                    LanguageConsts.TaskDueDateLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                    #endregion
                };

                if (message.Type == Storage.Message.MessageType.EmailEncryptedAndMaybeSigned ||
                    message.Type == Storage.Message.MessageType.EmailClearSigned)
                    languageConsts.Add(LanguageConsts.EmailSignedBy);

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] {0}).Max() + 2;
            }

            var emailHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel,
                message.GetEmailSender(htmlBody, hyperlinks));

            // Sent on
            if (message.SentOn != null)
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                    ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailRecipients(Storage.Recipient.RecipientType.To, htmlBody, hyperlinks));

            // CC
            var cc = message.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, htmlBody, hyperlinks);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailRecipients(Storage.Recipient.RecipientType.Bcc, htmlBody, hyperlinks);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            if (message.Type == Storage.Message.MessageType.EmailEncryptedAndMaybeSigned ||
                message.Type == Storage.Message.MessageType.EmailClearSigned)
            {
                var signerInfo = message.SignedBy;
                if (message.SignedOn != null)
                    signerInfo += " " + LanguageConsts.EmailSignedByOn + " " +
                                  ((DateTime) message.SignedOn).ToString(LanguageConsts.DataFormatWithTime);

                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSignedBy, signerInfo);
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
                    if (message.Task.Complete != null && (bool) message.Task.Complete)
                    {
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpStatusLabel,
                            LanguageConsts.EmailFollowUpCompletedText);

                        // Task completed date
                        if (message.Task.CompleteTime != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDateCompleted,
                                ((DateTime) message.Task.CompleteTime).ToString(LanguageConsts.DataFormatWithTime));
                    }
                    else
                    {
                        // Task startdate
                        if (message.Task.StartDate != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskStartDateLabel,
                                ((DateTime) message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

                        // Task duedate
                        if (message.Task.DueDate != null)
                            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDueDateLabel,
                                ((DateTime) message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));
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
        private static void SurroundWithHTML(StringBuilder footer,bool htmlBody)
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
        private static void WriteHeaderStart(StringBuilder header, bool htmlBody)
        {
            if (!htmlBody)
                return;

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
        private static void WriteHeaderLine(StringBuilder header,
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

                header.AppendLine(
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"font-weight: bold; white-space:nowrap;\">" +
                     WebUtility.HtmlEncode(label) + ":</td><td>" + newText + "</td></tr>");
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
        private static void WriteHeaderLineNoEncoding(StringBuilder header,
                                                      bool htmlBody,
                                                      int labelPadRightWidth,
                                                      string label,
                                                      string text)
        {
            if (htmlBody)
            {
                text = text.Replace("\n", "<br/>");

                header.AppendLine(
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"font-weight: bold; white-space:nowrap; \">" +
                    WebUtility.HtmlEncode(label) + ":</td><td>" + text + "</td></tr>");
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
        private static void WriteHeaderEmptyLine(StringBuilder header, bool htmlBody)
        {
            // Prevent that we write 2 empty lines in a row
            if (_emptyLineWritten)
                return;

            header.AppendLine(
                htmlBody
                    ? "<tr style=\"height: 18px; vertical-align: top; \"><td>&nbsp;</td><td>&nbsp;</td></tr>"
                    : string.Empty);

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
        /// <returns></returns>
        private List<string> WriteMsgEmail(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "email";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out htmlBody,
                out body,
                out dummy,
                out attachmentList,
                out files);

            var emailHeader = ExtractMsgEmailHeader(message, htmlBody, hyperlinks, attachmentList);
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
            var fileName = "email";
            bool htmlBody;
            string body;
            List<string> attachmentList;
            List<MemoryStream> attachStreams;
            List<MemoryStream> streams = new List<MemoryStream>();

            PreProcessEmlStream(message,
                hyperlinks,
                out htmlBody,
                out body,
                out attachmentList,
                out attachStreams);

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
                (message.Headers.DateSent.ToLocalTime()).ToString(LanguageConsts.DataFormatWithTime,
                    new CultureInfo(LanguageConsts.DataFormatWithTime)));

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

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            streams.Add(new MemoryStream(Encoding.UTF8.GetBytes(body)));

            /*******************************End Header*********************************/

            streams.AddRange(attachStreams);

            /*******************************Start Footer*******************************/
            var emailFooter = new StringBuilder();

            WriteHeaderStart(emailFooter, htmlBody);
            int i = 0;
            foreach (var item in headers.UnknownHeaders.AllKeys)
            {
                WriteHeaderLine(emailFooter, htmlBody, maxLength, item, headers.UnknownHeaders[i].ToString());
                i++;
            }
            SurroundWithHTML(emailFooter, htmlBody);
            streams.Add(new MemoryStream(Encoding.UTF8.GetBytes(emailFooter.ToString())));
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
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteEmlEmail(Mime.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "email";
            bool htmlBody;
            string body;
            List<string> attachmentList;
            List<string> files;

            PreProcessEmlFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out htmlBody,
                out body,
                out attachmentList,
                out files);

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
                (message.Headers.DateSent.ToLocalTime()).ToString(LanguageConsts.DataFormatWithTime));

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

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

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
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteMsgAppointment(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "appointment";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out htmlBody,
                out body,
                out dummy,
                out attachmentList,
                out files);

            if (!htmlBody)
                hyperlinks = false;

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

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] {0}).Max() + 2;
            }

            var appointmentHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(appointmentHeader, htmlBody);

            // Subject
            WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentSubjectLabel,
                message.Subject);

            // Location
            WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentLocationLabel,
                message.Appointment.Location);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Start
            if (message.Appointment.Start != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentStartDateLabel,
                    ((DateTime) message.Appointment.Start).ToString(LanguageConsts.DataFormatWithTime));

            // End
            if (message.Appointment.End != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentEndDateLabel,
                    ((DateTime) message.Appointment.End).ToString(LanguageConsts.DataFormatWithTime));

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Recurrence type
            if (!string.IsNullOrEmpty(message.Appointment.RecurrenceTypeText))
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrenceTypeLabel,
                    message.Appointment.RecurrenceTypeText);

            // Recurrence patern
            if (!string.IsNullOrEmpty(message.Appointment.RecurrencePatern))
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrencePaternLabel,
                    message.Appointment.RecurrencePatern);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Status
            if (message.Appointment.ClientIntentText != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentClientIntentLabel,
                    message.Appointment.ClientIntentText);

            // Appointment organizer (FROM)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentOrganizerLabel,
                message.GetEmailSender(htmlBody, hyperlinks));

            // Mandatory participants (TO)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                LanguageConsts.AppointmentMandatoryParticipantsLabel,
                message.GetEmailRecipients(Storage.Recipient.RecipientType.To, htmlBody, hyperlinks));

            // Optional participants (CC)
            var cc = message.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, htmlBody, hyperlinks);
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
                    String.Join("; ", categories));

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

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

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
        /// <param name="hyperlinks">When true then hyperlinks are generated attachments</param>
        /// <returns></returns>
        private List<string> WriteMsgTask(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "task";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out htmlBody,
                out body,
                out dummy,
                out attachmentList,
                out files);
            
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

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] {0}).Max() + 2;
            }

            var taskHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(taskHeader, htmlBody);

            // Subject
            WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskSubjectLabel, message.Subject);

            // Task startdate
            if (message.Task.StartDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskStartDateLabel,
                    ((DateTime) message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

            // Task duedate
            if (message.Task.DueDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskDueDateLabel,
                    ((DateTime) message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));

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
            if (message.Task.StatusText != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskStatusLabel, message.Task.StatusText);

            // Percentage complete
            if (message.Task.PercentageComplete != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskPercentageCompleteLabel,
                    (message.Task.PercentageComplete*100) + "%");

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // Estimated effort
            if (message.Task.EstimatedEffortText != null)
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
            if (message.Task.Owner != null)
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskOwnerLabel, message.Task.Owner);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Contacts
            if (message.Task.Contacts != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskContactsLabel,
                    string.Join("; ", message.Task.Contacts.ToArray()));

            // Categories
            var categories = message.Categories;
            if (categories != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

            // Companies
            if (message.Task.Companies != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskCompanyLabel,
                    string.Join("; ", message.Task.Companies.ToArray()));

            // Billing information
            if (message.Task.BillingInformation != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskBillingInformationLabel,
                    message.Task.BillingInformation);

            // Mileage
            if (message.Task.Mileage != null)
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

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

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
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteMsgContact(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "contact";
            bool htmlBody;
            string body;
            string contactPhotoFileName;
            List<string> attachmentList;
            List<string> files;

            PreProcessMsgFile(message,
                hyperlinks,
                outputFolder,
                ref fileName,
                out htmlBody,
                out body,
                out contactPhotoFileName,
                out attachmentList,
                out files);

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

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] {0}).Max() + 2;
            }

            var contactHeader = new StringBuilder();
            
            // Start of table
            WriteHeaderStart(contactHeader, htmlBody);

            if (htmlBody && !string.IsNullOrEmpty(contactPhotoFileName))
                contactHeader.Append(
                    "<div style=\"height: 250px; position: absolute; top: 20px; right: 20px;\"><img alt=\"\" src=\"" +
                    contactPhotoFileName + "\" height=\"100%\"></div>");

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
                    String.Join("; ", categories));

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            WriteHeaderEnd(contactHeader, htmlBody);
            
            body = InjectHeader(body, contactHeader.ToString());

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

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
        private static List<string> WriteMsgStickyNote(Storage.Message message, string outputFolder)
        {
            var files = new List<string>();
            string stickyNoteFile;
            var stickyNoteHeader = new StringBuilder();

            // Sticky notes only have RTF or Text bodies
            var body = message.BodyRtf;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
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
                                     : "stickynote") + ".txt";
            }

            // Write the body to a file
            stickyNoteFile = FileManager.FileExistsMakeNew(stickyNoteFile);
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            files.Add(stickyNoteFile);
            return files;
        }
        #endregion

        #region PreProcessMsgFile
        /// <summary>
        /// This method reads the body of a message object and returns it as an html body
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="htmlBody">Returns <c>true</c> when an html body is returned, <c>false</c>
        /// when the body is text based</param>
        /// <returns>True when the e-Mail has an HTML body</returns>
        private static string PreProcessMsgFile(Storage.Message message, out bool htmlBody)
        {
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
                    var converter = new RtfToHtmlConverter();
                    body = converter.ConvertRtfToHtml(body);
                    htmlBody = true;
                }
                else
                {
                    body = message.BodyText;

                    // When there is no body at all we just make an empty html document
                    if (body == null)
                    {
                        htmlBody = true;
                        body = "<html><head></head><body></body></html>";
                    }
                }
            }

            return body;
        }

        /// <summary>
        /// This function pre processes the Outlook MSG <see cref="Storage.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Storage.Attachment"/> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and 
        /// attachments (when there is an html body)</param>
        /// <param name="outputFolder">The outputfolder where alle extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Storage.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="contactPhotoFileName">Returns the filename of the contact photo. This field will only
        /// return a value when the <see cref="Storage.Message"/> object is a <see cref="Storage.Message.MessageType.Contact"/> 
        /// type and the <see cref="Storage.Message.Attachments"/> contains an object that has the 
        /// <param ref="Storage.Message.Attachment.IsContactPhoto"/> set to true, otherwise this field will always be null</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Storage.Message"/> object</param>
        private void PreProcessMsgFile(Storage.Message message,
            bool hyperlinks,
            string outputFolder,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out string contactPhotoFileName,
            out List<string> attachments,
            out List<string> files)
        {
            const string rtfInlineObject = "[*[RTFINLINEOBJECT]*]";

            htmlBody = true;
            attachments = new List<string>();
            files = new List<string>();
            contactPhotoFileName = null;
            body = message.BodyHtml;

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
                    var converter = new RtfToHtmlConverter();
                    body = converter.ConvertRtfToHtml(body);
                    htmlBody = true;
                }
                else
                {
                    body = message.BodyText;

                    // When there is no body at all we just make an empty html document
                    if (body == null)
                    {
                        htmlBody = true;
                        body = "<html><head></head><body></body></html>";
                    }
                }
            }

            fileName = outputFolder +
                       (!string.IsNullOrEmpty(message.Subject)
                           ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                           : fileName) + (htmlBody ? ".htm" : ".txt");

            fileName = FileManager.FileExistsMakeNew(fileName);
            files.Add(fileName);

            var inlineAttachments = new List<InlineAttachment>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;
                var attachmentFileName = string.Empty;
                var renderingPosition = -1;
                var isInline = false;

                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (attachment is Storage.Attachment)
                {
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
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
                        body = body.Replace("cid:" + attach.ContentId, fileInfo.FullName);
                    else
                        // If we didn't find the cid tag we treat the inline attachment as a normal one 
                        isInline = false;
                }
                // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
                else if (attachment is Storage.Message)
                // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
                {
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
                        using (var icon = FileIcon.GetFileIcon(fileInfo.FullName))
                        {
                            var iconFileName = outputFolder + Guid.NewGuid() + ".png";
                            icon.Save(iconFileName, ImageFormat.Png);
                            inlineAttachments.Add(new InlineAttachment(iconFileName, attachmentFileName, fileInfo.FullName));
                        }
                    else
                        inlineAttachments.Add(new InlineAttachment(renderingPosition, attachmentFileName));
                }
                else
                    renderingPosition = -1;

                if (!isInline && renderingPosition == -1)
                {
                    if (htmlBody)
                    {
                        if (hyperlinks)
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
            }

            if (htmlBody)
                foreach (var inlineAttachment in inlineAttachments.OrderBy(m => m.RenderingPosition))
                {
                    if (inlineAttachment.IconFileName != null)
                        body = ReplaceFirstOccurence(body, rtfInlineObject,
                            "<table style=\"width: 70px; display: inline; text-align: center; font-family: Times New Roman; font-size: 12pt;\"><tr><td>" +
                            (hyperlinks ? "<a href=\"" + inlineAttachment.FullName + "\">" : string.Empty) + "<img alt=\"\" src=\"" +
                            inlineAttachment.IconFileName + "\">" + (hyperlinks ? "</a>" : string.Empty) + "</td></tr><tr><td>" +
                            WebUtility.HtmlEncode(inlineAttachment.AttachmentFileName) +
                            "</td></tr></table>");
                    else
                        body = ReplaceFirstOccurence(body, rtfInlineObject, "<img alt=\"\" src=\"" + inlineAttachment.FullName + "\">");
                }
        }
        #endregion
        
        #region PreProcessEmlStream
        /// <summary>
        /// This function pre processes the EML <see cref="Mime.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Mime.MessagePart">attachment</see> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and
        /// attachments (when there is an html body)</param>
        /// <param name="htmlBody">Returns true when the <see cref="Mime.Message"/> object did contain
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Mime.Message"/> object</param>
        public void PreProcessEmlStream(Mime.Message message,
            bool hyperlinks,
            out bool htmlBody,
            out string body,
            out List<string> attachments,
            out List<MemoryStream> attachStreams)
        {
            attachments = new List<string>();
            attachStreams = new List<MemoryStream>();

            var bodyMessagePart = message.HtmlBody;

            if (bodyMessagePart != null)
            {
                body = bodyMessagePart.GetBodyAsText();
                htmlBody = true;
            }
            else
            {
                bodyMessagePart = message.TextBody;

                // When there is no body at all we just make an empty html document
                if (bodyMessagePart != null)
                {
                    body = bodyMessagePart.GetBodyAsText();
                    htmlBody = false;
                }
                else
                {
                    htmlBody = true;
                    body = "<html><head></head><body></body></html>";
                }
            }

            if (message.Attachments != null)
            {
                foreach (var attachment in message.Attachments)
                {
                    var attachmentFileName = attachment.FileName;

                    //use the stream here and don't worry about needing to close it
                    attachStreams.Add(new MemoryStream(attachment.Body));

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attachment.ContentId) && body.Contains(attachment.ContentId))
                    {
                        body = body.Replace("cid:" + attachment.ContentId, CheckValidAttachment(attachmentFileName));
                    }
                    else
                    {
                        // If we didn't find the cid tag we treat the inline attachment as a normal one

                        if (htmlBody)
                        {
                            if (hyperlinks)
                                attachments.Add("<a href=\"" + attachmentFileName + "\">" +
                                                HttpUtility.HtmlEncode(CheckValidAttachment(attachmentFileName)) + "</a> (" +
                                                FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                            else
                                attachments.Add(HttpUtility.HtmlEncode(CheckValidAttachment(attachmentFileName)) + " (" +
                                                FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                        }
                        else
                            attachments.Add(CheckValidAttachment(attachmentFileName) + " (" + FileManager.GetFileSizeString(attachment.Body.Length) + ")");
                    }
                }
            }
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
            string filename = attachmentFileName;
            string attchType = Path.GetExtension(attachmentFileName);
            switch (attchType)
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
        /// This function pre processes the EML <see cref="Mime.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Mime.MessagePart">attachment</see> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and 
        /// attachments (when there is an html body)</param>
        /// <param name="outputFolder">The outputfolder where alle extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Mime.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Mime.Message"/> object</param>
        private static void PreProcessEmlFile(Mime.Message message,
            bool hyperlinks,
            string outputFolder,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out List<string> attachments,
            out List<string> files)
        {
            attachments = new List<string>();
            files = new List<string>();

            var bodyMessagePart = message.HtmlBody;

            if (bodyMessagePart != null)
            {
                body = bodyMessagePart.GetBodyAsText();
                htmlBody = true;
            }
            else
            {
                bodyMessagePart = message.TextBody;

                // When there is no body at all we just make an empty html document
                if (bodyMessagePart != null)
                {
                    body = bodyMessagePart.GetBodyAsText();
                    htmlBody = false;
                }
                else
                {
                    htmlBody = true;
                    body = "<html><head></head><body></body></html>";
                }
            }

            fileName = outputFolder +
                       (!string.IsNullOrEmpty(message.Headers.Subject)
                           ? FileManager.RemoveInvalidFileNameChars(message.Headers.Subject)
                           : fileName) + (htmlBody ? ".htm" : ".txt");

            fileName = FileManager.FileExistsMakeNew(fileName);
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
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attachment.ContentId) && body.Contains(attachment.ContentId))
                    {
                        body = body.Replace("cid:" + attachment.ContentId, fileInfo.FullName);
                    }
                    else
                    {
                        // If we didn't find the cid tag we treat the inline attachment as a normal one 

                        files.Add(fileInfo.FullName);

                        if (htmlBody)
                        {
                            if (hyperlinks)
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
                }
            }
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
        private static string InjectHeader(string body, string header, string contentType = null)
        {
            var begin = body.IndexOf("<BODY", StringComparison.InvariantCultureIgnoreCase);

            if (begin <= 0) return header + body;
            begin = body.IndexOf(">", begin, StringComparison.InvariantCultureIgnoreCase);

            body = body.Insert(begin + 1, header);

            if (!string.IsNullOrWhiteSpace(contentType))
            {
                // Inject content-type:
                var head = "<head";
                var headBegin = body.IndexOf(head, StringComparison.InvariantCultureIgnoreCase) + head.Length;
                headBegin = body.IndexOf(">", headBegin, StringComparison.InvariantCultureIgnoreCase);

                var contentHeader = string.Format("{0}<meta http-equiv=\"Content-Type\" content=\"{1}\">{2}", Environment.NewLine,
                    contentType, Environment.NewLine);

                body = body.Insert(headBegin + 1, contentHeader);
            }

            return body;
        }
        #endregion
    }
}
