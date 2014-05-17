using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DocumentServices.Modules.Readers.MsgReader.Helpers;
using DocumentServices.Modules.Readers.MsgReader.Outlook;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace DocumentServices.Modules.Readers.MsgReader
{
    #region Interface IReader
    public interface IReader
    {
        /// <summary>
        /// Extract the input msg file to the given output folder
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to extract the msg file</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns>String array containing the message body and its (inline) attachments</returns>
        [DispId(1)]
        string[] ExtractToFolder(string inputFile, string outputFolder, bool hyperlinks = false);

        /// <summary>
        /// This function will read all the properties of an <see cref="Storage.Message"/> file and maps
        /// all the properties that are filled to the extended file attributes. 
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        [DispId(2)]
        void SetExtendedFileAttributesWithMsgProperties(string inputFile);

        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(3)]
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
        /// Contains an error message when something goes wrong in the <see cref="ExtractToFolder"/> method.
        /// This message can be retreived with the GetErrorMessage. This way we keep .NET exceptions inside
        /// when this code is called from a COM language
        /// </summary>
        private string _errorMessage;

        /// <summary>
        /// Used to keep track if we already did write an empty line
        /// </summary>
        private static bool _emptyLineWritten;
        #endregion

        #region Private nested class Recipient
        /// <summary>
        /// Used as a placeholder for the recipients from the MSG file itself or from the "internet"
        /// headers when this message is send outside an Exchange system
        /// </summary>
        private class Recipient
        {
            public string EmailAddress { get; set; }
            public string DisplayName { get; set; }
        }
        #endregion

        #region ExtractToFolder
        /// <summary>
        /// This method reads the <see cref="inputFile"/> and when the file is supported it will do the following: <br/>
        /// - Extract the HTML, RTF (will be converted to html) or TEXT body (in these order) <br/>
        /// - Puts a header (with the sender, to, cc, etc... (depends on the message type) on top of the body so it looks 
        ///   like if the object is printed from Outlook <br/>
        /// - Reads all the attachents <br/>
        /// And in the end writes everything to the given <see cref="outputFolder"/>
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to save the extracted msg file</param>
        /// <param name="hyperlinks">When true hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns>String array containing the full path to the message body and its attachments</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")]
        public string[] ExtractToFolder(string inputFile, string outputFolder, bool hyperlinks = false)
        {
            outputFolder = FileManager.CheckForBackSlash(outputFolder);
            _errorMessage = string.Empty;

            try
            {
                using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                using (var message = new Storage.Message(stream))
                {
                    switch (message.Type)
                    {
                        case Storage.Message.MessageType.Email:
                            return WriteEmail(message, outputFolder, hyperlinks).ToArray();

                        case Storage.Message.MessageType.AppointmentRequest:
                        case Storage.Message.MessageType.Appointment:
                        case Storage.Message.MessageType.AppointmentResponse:
                            return WriteAppointment(message, outputFolder, hyperlinks).ToArray();

                        case Storage.Message.MessageType.Task:
                        case Storage.Message.MessageType.TaskRequestAccept:
                            return WriteTask(message, outputFolder, hyperlinks).ToArray();

                        case Storage.Message.MessageType.Contact:
                            return WriteContact(message, outputFolder, hyperlinks).ToArray();

                        case Storage.Message.MessageType.StickyNote:
                            return WriteStickyNote(message, outputFolder).ToArray();

                        case Storage.Message.MessageType.Unknown:
                            throw new NotSupportedException("Unknown message type");
                    }
                }
            }
            catch (Exception e)
            {
                _errorMessage = ExceptionHelpers.GetInnerException(e);
                return new string[0];
            }

            // If we return here then the file was not supported
            return new string[0];
        }
        #endregion

        #region ReplaceFirstOccurence
        /// <summary>
        /// Method to replace the first occurence of the <see cref="search"/> string with a
        /// <see cref="replace"/> string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="search"></param>
        /// <param name="replace"></param>
        /// <returns></returns>
        public string ReplaceFirstOccurence(string text, string search, string replace)
        {
            var index = text.IndexOf(search, StringComparison.Ordinal);
            if (index < 0)
                return text;

            return text.Substring(0, index) + replace + text.Substring(index + search.Length);
        }
        #endregion

        #region WriteHeader methods
        /// <summary>
        /// Writes the start of the header
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <see cref="header"/> otherwise text will be written</param>
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
        /// <param name="htmlBody">When true then html will be written into the <see cref="header"/> otherwise text will be written</param>
        /// <param name="labelPadRightWidth">Used to pad the label size, ignored when <see cref="htmlBody"/> is true</param>
        /// <param name="label">The label text that needs to be written</param>
        /// <param name="text">The text that needs to be written after the <see cref="label"/></param>
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
                    newText += HttpUtility.HtmlEncode(line) + "<br/>";

                header.AppendLine(
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"font-weight: bold; white-space:nowrap;\">" +
                    HttpUtility.HtmlEncode(label) + ":</td><td>" + newText + "</td></tr>");
            }
            else
                        {
                text = text.Replace("\n", "".PadRight(labelPadRightWidth));
                header.AppendLine((label + ":").PadRight(labelPadRightWidth) + text);
            }

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes a line into the header without Html encoding the <see cref="text"/>
        /// </summary>
        /// <param name="header">The <see cref="StringBuilder"/> object that is used to write a header</param>
        /// <param name="htmlBody">When true then html will be written into the <see cref="header"/> otherwise text will be written</param>
        /// <param name="labelPadRightWidth">Used to pad the label size, ignored when <see cref="htmlBody"/> is true</param>
        /// <param name="label">The label text that needs to be written</param>
        /// <param name="text">The text that needs to be written after the <see cref="label"/></param>
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
                    HttpUtility.HtmlEncode(label) + ":</td><td>" + text + "</td></tr>");
            }
            else
            {
                text = text.Replace("\n", "".PadRight(labelPadRightWidth));
                header.AppendLine((label + ":").PadRight(labelPadRightWidth) + text);
            }

            _emptyLineWritten = false;
        }

        /// <summary>
        /// Writes
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
        /// <param name="htmlBody">When true then html will be written into the <see cref="header"/> otherwise text will be written</param>
        private static void WriteHeaderEnd(StringBuilder header, bool htmlBody)
        {
            if (!htmlBody)
                return;

            header.AppendLine("</table><br/>");
        }
        #endregion

        #region WriteEmail
        /// <summary>
        /// Writes the body of the MSG E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteEmail(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "email";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMesssage(message,
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

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] {0}).Max() + 2;
            }
            
            var emailHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel,
                GetEmailSender(message, hyperlinks, htmlBody));

            // Sent on
            if (message.SentOn != null)
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                    ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, htmlBody));

            // CC
            var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

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
            if (attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                    string.Join(", ", attachmentList));

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // Follow up
            if (message.Flag != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpLabel,
                    message.Flag.Request);

                // When complete
                if (message.Task.Complete != null && (bool) message.Task.Complete)
                {
                    WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpStatusLabel,
                        LanguageConsts.EmailFollowUpCompletedText);

                    // Task completed date
                    if (message.Task.CompleteTime != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDateCompleted,
                            ((DateTime) message.Task.CompleteTime).ToString(LanguageConsts.DataFormatWithTime,
                                new CultureInfo(LanguageConsts.DateFormatCulture)));
                }
                else
                {
                    // Task startdate
                    if (message.Task.StartDate != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskStartDateLabel,
                            ((DateTime) message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime,
                                new CultureInfo(LanguageConsts.DateFormatCulture)));

                    // Task duedate
                    if (message.Task.DueDate != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDueDateLabel,
                            ((DateTime) message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime,
                                new CultureInfo(LanguageConsts.DateFormatCulture)));
                }

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Categories
            var categories = message.Categories;
            if (categories != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            body = InjectHeader(body, emailHeader.ToString());

            // Write the body to a file
            File.WriteAllText(fileName, body, Encoding.UTF8);

            return files;
        }
        #endregion

        #region MapEmailPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Message"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapEmailPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            if (message.SentOn != null)
                propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);
            
            // To
            propertyWriter.WriteProperty(SystemProperties.System.Message.ToAddress, GetEmailRecipients(message, Storage.Recipient.RecipientType.To, false, false));
            
            // CC
            var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, false, false);
            if (!string.IsNullOrEmpty(cc))
                propertyWriter.WriteProperty(SystemProperties.System.Message.CcAddress, cc);
            
            // BCC
            var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, false, false);
            if (!string.IsNullOrEmpty(bcc))
                propertyWriter.WriteProperty(SystemProperties.System.Message.BccAddress, bcc);

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);
            
            // Urgent
            propertyWriter.WriteProperty(SystemProperties.System.Importance, message.Importance);
            propertyWriter.WriteProperty(SystemProperties.System.ImportanceText, message.ImportanceText);
     
            // Attachments
            var attachments = GetAttachmentNames(message);
            if (string.IsNullOrEmpty(attachments))
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, false);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, null);
            }
            else
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, true);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, attachments);
            }

            // Clear properties
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, null);
            propertyWriter.WriteProperty(SystemProperties.System.DueDate, null);
            propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, null);
            propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, null);
            propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, null);

            // Follow up
            if (message.Flag != null)
            {
                propertyWriter.WriteProperty(SystemProperties.System.IsFlagged, true);
                propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, message.Flag.Request);
                
                // Flag status text
                propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, message.Task.StatusText);

                // When complete
                if (message.Task.Complete != null && (bool)message.Task.Complete)
                {
                    // Flagged complete
                    propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, true);
                    
                    // Task completed date
                    if (message.Task.CompleteTime != null)
                        propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, (DateTime)message.Task.CompleteTime);
                }
                else
                {
                    // Flagged not complete
                    propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, false);
                    
                    propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, null);

                    // Task startdate
                    if (message.Task.StartDate != null)
                        propertyWriter.WriteProperty(SystemProperties.System.StartDate, (DateTime)message.Task.StartDate);

                    // Task duedate
                    if (message.Task.DueDate != null)
                        propertyWriter.WriteProperty(SystemProperties.System.DueDate, (DateTime)message.Task.DueDate);
                }
            }

            // Categories
            var categories = message.Categories;
            if (categories != null)
                propertyWriter.WriteProperty(SystemProperties.System.Category, String.Join("; ", String.Join("; ", categories)));
        }
        #endregion

        #region WriteAppointment
        /// <summary>
        /// Writes the body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteAppointment(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "appointment";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMesssage(message,
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
                    ((DateTime) message.Appointment.Start).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // End
            if (message.Appointment.End != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentEndDateLabel,
                    ((DateTime) message.Appointment.End).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Recurrence type
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
                GetEmailSender(message, hyperlinks, htmlBody));

            // Mandatory participants (TO)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                LanguageConsts.AppointmentMandatoryParticipantsLabel,
                GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, htmlBody));

            // Optional participants (CC)
            var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, htmlBody);
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

        #region MapAppointmentPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Appointment"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapAppointmentPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            if (message.SentOn != null)
                propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);

            // Location
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.Location, message.Appointment.Location);

            // Start
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, message.Appointment.Start);

            // End
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, message.Appointment.End);

            // Recurrence type
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.IsRecurring,
                message.Appointment.ReccurrenceType != Storage.Appointment.AppointmentRecurrenceType.None);

            // Status
            propertyWriter.WriteProperty(SystemProperties.System.Status, message.Appointment.ClientIntentText);

            // Appointment organizer (FROM)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OrganizerAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OrganizerName, message.Sender.DisplayName);

            // Mandatory participants (TO)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.RequiredAttendeeNames, message.Appointment.ToAttendees);

            // Optional participants (CC)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OptionalAttendeeNames, message.Appointment.CclAttendees);

            // Categories
            var categories = message.Categories;
            if (categories != null)
                propertyWriter.WriteProperty(SystemProperties.System.Category, String.Join("; ", String.Join("; ", categories)));

            // Urgent
            propertyWriter.WriteProperty(SystemProperties.System.Importance, message.Importance);
            propertyWriter.WriteProperty(SystemProperties.System.ImportanceText, message.ImportanceText);

            // Attachments
            var attachments = GetAttachmentNames(message);
            if (string.IsNullOrEmpty(attachments))
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, false);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, null);
            }
            else
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, true);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, attachments);
            }
        }
        #endregion

        #region WriteTask
        /// <summary>
        /// Writes the task body of the MSG Task to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated attachments</param>
        /// <returns></returns>
        private List<string> WriteTask(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "task";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMesssage(message,
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
                    ((DateTime) message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // Task duedate
            if (message.Task.DueDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskDueDateLabel,
                    ((DateTime) message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

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

        #region MapTaskPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Task"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapTaskPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            if (message.SentOn != null)
                propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);

            /*
            // Subject
            WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskSubjectLabel, message.Subject);

            // Task startdate
            if (message.Task.StartDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskStartDateLabel,
                    ((DateTime) message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // Task duedate
            if (message.Task.DueDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskDueDateLabel,
                    ((DateTime) message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // Urgent
            var importance = message.ImportanceText;
            if (importance != null)
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
            */
        }
        #endregion

        #region WriteContact
        /// <summary>
        /// Writes the body of the MSG Contact to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteContact(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            var fileName = "contact";
            bool htmlBody;
            string body;
            string contactPhotoFileName;
            List<string> attachmentList;
            List<string> files;

            PreProcessMesssage(message,
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
                    ((DateTime)message.Contact.Birthday).ToString(LanguageConsts.DataFormat,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

            // Anniversary
            if (message.Contact.WeddingAnniversary != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.WeddingAnniversaryLabel,
                    ((DateTime)message.Contact.WeddingAnniversary).ToString(LanguageConsts.DataFormat,
                        new CultureInfo(LanguageConsts.DateFormatCulture)));

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

        #region WriteStickyNote
        /// <summary>
        /// Writes the body of the MSG StickyNote to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <returns></returns>
        private List<string> WriteStickyNote(Storage.Message message, string outputFolder)
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
                        ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

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
                        ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

                body = stickyNoteHeader + body;
                stickyNoteFile = outputFolder +
                                 (!string.IsNullOrEmpty(message.Subject)
                                     ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                     : "stickynote") + ".txt";
            }

            // Write the body to a file
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            files.Add(stickyNoteFile);
            return files;
        }
        #endregion

        #region GetAttachmentNames
        /// <summary>
        /// Returns the attachments names as a comma sperated string
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <returns></returns>
        private string GetAttachmentNames(Storage.Message message)
        {
            var result = new List<string>();

            foreach (var attachment in message.Attachments)
            {
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    result.Add(attach.FileName);
                }
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;
                    result.Add(msg.FileName);
                }
            }

            return string.Join(", ", result);
        }
        #endregion

        #region PreProcessMesssage
        /// <summary>
        /// This function pre processes the <see cref="Storage.Message"/> object, it tries to find the html (or text) body
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
        /// <see cref="Storage.Message.Attachment.IsContactPhoto"/> set to true, otherwise this field will always be null</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Storage.Message"/> object</param>
        private void PreProcessMesssage(Storage.Message message,
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
            var htmlConvertedFromRtf = false;
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
                    htmlConvertedFromRtf = true;
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

            files.Add(fileName);

            var inlineAttachments = new SortedDictionary<int, string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;
                var attachmentFileName = string.Empty;
                var renderingPosition = -1;
                var isInline = false;

                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment) attachment;
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

                    if (!htmlConvertedFromRtf)
                    {
                        // When we find an inline attachment we have to replace the CID tag inside the html body
                        // with the name of the inline attachment. But before we do this we check if the CID exists.
                        // When the CID does not exists we treat the inline attachment as a normal attachment
                        if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
                            body = body.Replace("cid:" + attach.ContentId, fileInfo.FullName);
                        else
                            // If we didn't find the cid tag we treat the inline attachment as a normal one 
                            isInline = false;
                    }
                }
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message) attachment;
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
                if (htmlBody && renderingPosition != -1)
                {
                    if (!isInline)
                        using (var icon = FileIcon.GetFileIcon(fileInfo.FullName))
                        {
                            var iconFileName = outputFolder + Guid.NewGuid() + ".png";
                            icon.Save(iconFileName, ImageFormat.Png);
                            inlineAttachments.Add(renderingPosition,
                                iconFileName + "|" + attachmentFileName + "|" + fileInfo.FullName);
                        }
                    else
                        inlineAttachments.Add(renderingPosition, attachmentFileName);
                }
                else
                    renderingPosition = -1;

                if (!isInline && renderingPosition == -1)
                {
                    if (htmlBody)
                    {
                        if (hyperlinks)
                            attachments.Add("<a href=\"" + fileInfo.Name + "\">" +
                                            HttpUtility.HtmlEncode(attachmentFileName) + "</a> (" +
                                            FileManager.GetFileSizeString(fileInfo.Length) + ")");
                        else
                            attachments.Add(HttpUtility.HtmlEncode(attachmentFileName) + " (" +
                                            FileManager.GetFileSizeString(fileInfo.Length) + ")");
                    }
                    else
                        attachments.Add(attachmentFileName + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
                }
            }

            if (htmlBody)
                foreach (var inlineAttachment in inlineAttachments)
                {
                    var names = inlineAttachment.Value.Split('|');

                    if (names.Length == 3)
                        body = ReplaceFirstOccurence(body, rtfInlineObject,
                            "<table style=\"width: 70px; display: inline; text-align: center; font-family: Times New Roman; font-size: 12pt;\"><tr><td>" +
                            (hyperlinks ? "<a href=\"" + names[2] + "\">" : string.Empty) + "<img alt=\"\" src=\"" +
                            names[0] + "\">" + (hyperlinks ? "</a>" : string.Empty) + "</td></tr><tr><td>" +
                            HttpUtility.HtmlEncode(names[1]) +
                            "</td></tr></table>");
                    else
                        body = ReplaceFirstOccurence(body, rtfInlineObject, "<img alt=\"\" src=\"" + names[0] + "\">");
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

        #region RemoveSingleQuotes
        /// <summary>
        /// Removes trailing en ending single quotes from an E-mail address when they exist
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        private static string RemoveSingleQuotes(string email)
        {
            if (string.IsNullOrEmpty(email))
                return string.Empty;

            if (email.StartsWith("'"))
                email = email.Substring(1, email.Length - 1);

            if (email.EndsWith("'"))
                email = email.Substring(0, email.Length - 1);

            return email;
        }
        #endregion

        #region IsEmailAddressValid
        /// <summary>
        /// Return true when the E-mail address is valid
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        private static bool IsEmailAddressValid(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress))
                return false;

            var regex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
            var matches = regex.Matches(emailAddress);

            return matches.Count == 1;
        }
        #endregion

        #region GetEmailSender
        /// <summary>
        /// Changes the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true then E-mail addresses are converted to hyperlinks</param>
        /// <param name="html">Set this to true when the E-mail body format is html</param>
        /// <returns></returns>
        private static string GetEmailSender(Storage.Message message, bool convertToHref, bool html)
        {
            var output = string.Empty;

            if (message == null) return string.Empty;

            var tempEmailAddress = message.Sender.Email;
            var tempDisplayName = message.Sender.DisplayName;

            if (string.IsNullOrEmpty(tempEmailAddress) && message.Headers != null && message.Headers.From != null)
                tempEmailAddress = RemoveSingleQuotes(message.Headers.From.Address);

            if (string.IsNullOrEmpty(tempDisplayName) && message.Headers != null && message.Headers.From != null)
                tempDisplayName = message.Headers.From.DisplayName;

            var emailAddress = tempEmailAddress;
            var displayName = tempDisplayName;

            // Sometimes the E-mail address and displayname get swapped so check if they are valid
            if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
            {
                // Swap them
                emailAddress = tempDisplayName;
                displayName = tempEmailAddress;
            }
            else if (IsEmailAddressValid(tempDisplayName))
            {
                // If the displayname is an emailAddress them move it
                emailAddress = tempDisplayName;
                displayName = tempDisplayName;
            }

            if (string.Equals(emailAddress, displayName, StringComparison.InvariantCultureIgnoreCase))
                displayName = string.Empty;

            if (html)
            {
                emailAddress = HttpUtility.HtmlEncode(emailAddress);
                displayName = HttpUtility.HtmlEncode(displayName);
            }

            if (convertToHref && html && !string.IsNullOrEmpty(emailAddress))
                output += "<a href=\"mailto:" + emailAddress + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                              ? displayName
                              : emailAddress) + "</a>";

            else
            {
                if (!string.IsNullOrEmpty(displayName))
                    output = displayName;

                var beginTag = string.Empty;
                var endTag = string.Empty;
                if (!string.IsNullOrEmpty(displayName))
                {
                    if (html)
                    {
                        beginTag = "&nbsp&lt;";
                        endTag = "&gt;";
                    }
                    else
                    {
                        beginTag = " <";
                        endTag = ">";
                    }
                }

                if (!string.IsNullOrEmpty(emailAddress))
                    output += beginTag + emailAddress + endTag;
            }

            return output;
        }
        #endregion

        #region GetEmailRecipients
        /// <summary>
        /// Change the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <param name="type">Selects the Recipient type to retrieve</param>
        /// <param name="html">Set this to true when the E-mail body format is html</param>
        /// <returns></returns>
        private static string GetEmailRecipients(Storage.Message message,
            Storage.Recipient.RecipientType type,
            bool convertToHref,
            bool html)
        {
            var output = string.Empty;

            var recipients = new List<Recipient>();

            if (message == null)
                return output;

            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (var recipient in message.Recipients)
            {
                // First we filter for the correct recipient type
                if (recipient.Type == type)
                    recipients.Add(new Recipient {EmailAddress = recipient.Email, DisplayName = recipient.DisplayName});
            }

            // TODO move this code to the recipient class
            if (recipients.Count == 0 && message.Headers != null)
            {
                switch (type)
                {
                    case Storage.Recipient.RecipientType.To:
                        if (message.Headers.To != null)
                            recipients.AddRange(
                                message.Headers.To.Select(
                                    to => new Recipient {EmailAddress = to.Address, DisplayName = to.DisplayName}));
                        break;

                    case Storage.Recipient.RecipientType.Cc:
                        if (message.Headers.Cc != null)
                            recipients.AddRange(
                                message.Headers.Cc.Select(
                                    cc => new Recipient {EmailAddress = cc.Address, DisplayName = cc.DisplayName}));
                        break;

                    case Storage.Recipient.RecipientType.Bcc:
                        if (message.Headers.Bcc != null)
                            recipients.AddRange(
                                message.Headers.Bcc.Select(
                                    bcc => new Recipient {EmailAddress = bcc.Address, DisplayName = bcc.DisplayName}));
                        break;
                }
            }

            foreach (var recipient in recipients)
            {
                if (output != string.Empty)
                    output += "; ";

                var tempEmailAddress = RemoveSingleQuotes(recipient.EmailAddress);
                var tempDisplayName = RemoveSingleQuotes(recipient.DisplayName);

                var emailAddress = tempEmailAddress;
                var displayName = tempDisplayName;

                // Sometimes the E-mail address and displayname get swapped so check if they are valid
                if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
                {
                    // Swap them
                    emailAddress = tempDisplayName;
                    displayName = tempEmailAddress;
                }
                else if (IsEmailAddressValid(tempDisplayName))
                {
                    // If the displayname is an emailAddress them move it
                    emailAddress = tempDisplayName;
                    displayName = tempDisplayName;
                }

                if (string.Equals(emailAddress, displayName, StringComparison.InvariantCultureIgnoreCase))
                    displayName = string.Empty;

                if (html)
                {
                    emailAddress = HttpUtility.HtmlEncode(emailAddress);
                    displayName = HttpUtility.HtmlEncode(displayName);
                }

                if (convertToHref && html && !string.IsNullOrEmpty(emailAddress))
                    output += "<a href=\"mailto:" + emailAddress + "\">" +
                              (!string.IsNullOrEmpty(displayName)
                                  ? displayName
                                  : emailAddress) + "</a>";

                else
                {
                    if (!string.IsNullOrEmpty(displayName))
                        output += displayName;

                    var beginTag = string.Empty;
                    var endTag = string.Empty;
                    if (!string.IsNullOrEmpty(displayName))
                    {
                        if (html)
                        {
                            beginTag = "&nbsp&lt;";
                            endTag = "&gt;";
                        }
                        else
                        {
                            beginTag = " <";
                            endTag = ">";
                        }
                    }

                    if (!string.IsNullOrEmpty(emailAddress))
                        output += beginTag + emailAddress + endTag;
                }
            }

            return output;
        }
        #endregion

        #region InjectHeader
        /// <summary>
        /// Inject an outlook style header into the top of the html
        /// </summary>
        /// <param name="body"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        private static string InjectHeader(string body, string header)
        {
            var temp = body.ToUpperInvariant();

            var begin = temp.IndexOf("<BODY", StringComparison.Ordinal);

            if (begin > 0)
            {
                begin = temp.IndexOf(">", begin, StringComparison.Ordinal);
                return body.Insert(begin + 1, header);
            }

            return header + body;
        }
        #endregion

        #region SetExtendedFileAttributesWithMsgProperties
        /// <summary>
        /// This function will read all the properties of an <see cref="Storage.Message"/> file and maps
        /// all the properties that are filled to the extended file attributes. 
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        public void SetExtendedFileAttributesWithMsgProperties(string inputFile)
        {
            MemoryStream memoryStream = null;

            try
            {
                // We need to read the msg file into memory because we otherwise can't set the extended filesystem
                // properties because the files is locked
                memoryStream = new MemoryStream();
                using (var fileStream = File.OpenRead(inputFile))
                    fileStream.CopyTo(memoryStream);

                memoryStream.Position = 0;

                using (var shellFile = ShellFile.FromFilePath(inputFile))
                {
                    using (var propertyWriter = shellFile.Properties.GetPropertyWriter())
                    {
                        using (var message = new Storage.Message(memoryStream))
                        {
                            switch (message.Type)
                            {
                                case Storage.Message.MessageType.Email:
                                    MapEmailPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case Storage.Message.MessageType.AppointmentRequest:
                                case Storage.Message.MessageType.Appointment:
                                case Storage.Message.MessageType.AppointmentResponse:
                                    MapAppointmentPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case Storage.Message.MessageType.Task:
                                case Storage.Message.MessageType.TaskRequestAccept:
                                    MapTaskPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case Storage.Message.MessageType.Contact:
                                    //return WriteContact(message, outputFolder, hyperlinks).ToArray();
                                    break;

                                case Storage.Message.MessageType.StickyNote:
                                    //return WriteStickyNote(message, outputFolder).ToArray();
                                    break;

                                case Storage.Message.MessageType.Unknown:
                                    throw new NotSupportedException("Unsupported message type");
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                _errorMessage = ExceptionHelpers.GetInnerException(e);
            }
            finally
            {
                if (memoryStream != null)
                    memoryStream.Dispose();
            }
        }
        #endregion
    }
}