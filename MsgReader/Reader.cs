using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DocumentServices.Modules.Readers.MsgReader.Outlook;

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
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        string GetErrorMessage();
    }
    #endregion

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
        /// Extract the input msg file to the given output folder
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to extract the msg file</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns>String array containing the message body and its (inline) attachments</returns>
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
                            throw new Exception("An task file is not yet supported");

                        case Storage.Message.MessageType.Contact:
                            throw new Exception("An contact file is not yet supported");

                        case Storage.Message.MessageType.StickyNote:
                            return WriteStickyNote(message, outputFolder).ToArray();

                        case Storage.Message.MessageType.Unknown:
                            throw new NotSupportedException("Unknown message type");
                    }
                }
            }
            catch (Exception e)
            {
                _errorMessage = GetInnerException(e);
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
            var result = new List<string>();

            // We first always check if there is a HTML body
            var body = message.BodyHtml;
            var htmlBody = true;

            if (body == null)
            {
                // When there is not HTML body found then try to get the text body
                body = message.BodyText;
                htmlBody = false;
            }

            // Determine the name for the E-mail body
            var eMailFileName = outputFolder +
                                (!string.IsNullOrEmpty(message.Subject)
                                    ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                    : "email") + (htmlBody ? ".htm" : ".txt");

            result.Add(eMailFileName);

            #region Attachments
            var attachmentList = new List<string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;

                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attach.FileName));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody)
                    {
                        if (!string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
                        {
                            body = body.Replace("cid:" + attach.ContentId, fileInfo.FullName);
                            continue;
                        }
                    }

                    result.Add(fileInfo.FullName);
                }
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;

                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + msg.FileName) + ".msg");
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);
                }

                if (fileInfo == null) continue;

                if (htmlBody && hyperlinks)
                    attachmentList.Add("<a href=\"" + HttpUtility.HtmlEncode(fileInfo.Name) + "\">" +
                                       HttpUtility.HtmlEncode(fileInfo.Name) + "</a> (" +
                                       FileManager.GetFileSizeString(fileInfo.Length) + ")");
                else
                    attachmentList.Add(fileInfo.Name + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
            }
            #endregion

            string emailHeader;

            if (htmlBody)
            {
                #region Html body
                // Start of table
                emailHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // From
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.EmailFromLabel + ":</td><td>" + GetEmailSender(message, hyperlinks, true) +
                    "</td></tr>" + Environment.NewLine;

                // Sent on
                if (message.SentOn != null)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailSentOnLabel + ":</td><td>" +
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat, new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                        Environment.NewLine;

                // To
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.EmailToLabel + ":</td><td>" +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, true) + "</td></tr>" +
                    Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCcLabel + ":</td><td>" + cc + "</td></tr>" + Environment.NewLine;

                // BCC
                var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, hyperlinks, false);
                if (bcc != string.Empty)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailBccLabel + ":</td><td>" + bcc + "</td></tr>" + Environment.NewLine;

                // Subject
                emailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.EmailSubjectLabel + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.ImportanceLabel + ":</td><td>" + importance + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailAttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList) + "</td></tr>" +
                        Environment.NewLine;

                // Empty line
                emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Follow up
                if (message.Flag != null)
                {
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailFollowUpLabel + ":</td><td>" + message.Flag.Request + "</td></tr>" + Environment.NewLine;

                    // When complete
                    if (message.Task.Complete != null && (bool)message.Task.Complete)
                    {
                        emailHeader +=
                            "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                            LanguageConsts.EmailFollowUpStatusLabel + ":</td><td>" + LanguageConsts.EmailFollowUpCompletedText +
                            "</td></tr>" + Environment.NewLine;

                        // Task completed date
                        var completedDate = message.Task.CompleteTime;
                        if (completedDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.TaskDateCompleted + ":</td><td>" +
                                ((DateTime) completedDate).ToString(LanguageConsts.DataFormat,
                                    new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                                Environment.NewLine;
                    }
                    else
                    {
                        // Task startdate
                        var startDate = message.Task.StartDate;
                        if (startDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.TaskStartDateLabel + ":</td><td>" +
                                ((DateTime) startDate).ToString(LanguageConsts.DataFormat,
                                    new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                                Environment.NewLine;

                        // Task duedate
                        var dueDate = message.Task.DueDate;
                        if (dueDate != null)
                            emailHeader +=
                                "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                                LanguageConsts.TaskDueDateLabel + ":</td><td>" +
                                ((DateTime) dueDate).ToString(LanguageConsts.DataFormat,
                                    new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                                Environment.NewLine;

                    }

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    emailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories) + "</td></tr>" +
                        Environment.NewLine;

                    // Empty line
                    emailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // End of table + empty line
                emailHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, emailHeader);
                #endregion
            }
            else
            {
                #region Text body
                // Read all the language consts and get the longest string
                var languageConsts = new List<string>
                {
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
                };

                var maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max();

                // From
                emailHeader =
                    (LanguageConsts.EmailFromLabel + ":").PadRight(maxLength) + GetEmailSender(message, false, false) + Environment.NewLine;

                // Sent on
                if (message.SentOn != null)
                    emailHeader +=
                        (LanguageConsts.EmailSentOnLabel + ":").PadRight(maxLength) +
                        ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormat, new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine;

                // To
                emailHeader +=
                    (LanguageConsts.EmailToLabel + ":").PadRight(maxLength) +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, false, false) + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, false, false);
                if (cc != string.Empty)
                    emailHeader += (LanguageConsts.EmailCcLabel + ":").PadRight(maxLength) + cc + Environment.NewLine;

                // BCC
                var bcc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Bcc, false, false);
                if (bcc != string.Empty)
                    emailHeader += (LanguageConsts.EmailCcLabel + ":").PadRight(maxLength) + bcc + Environment.NewLine;

                // Subject
                emailHeader += (LanguageConsts.EmailSubjectLabel + ":").PadRight(maxLength) + message.Subject + Environment.NewLine;

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    // Importance text + new line
                    emailHeader += (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) + importance + Environment.NewLine + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                    emailHeader += (LanguageConsts.EmailAttachmentsLabel + ":").PadRight(maxLength) +
                                          string.Join(", ", attachmentList) + Environment.NewLine + Environment.NewLine;

                // Follow up
                if (message.Flag != null)
                {
                    emailHeader += (LanguageConsts.EmailFollowUpLabel + ":").PadRight(maxLength) + message.Flag.Request + Environment.NewLine;

                    // When complete
                    if (message.Task.Complete != null && (bool)message.Task.Complete)
                    {
                        emailHeader += (LanguageConsts.EmailFollowUpStatusLabel + ":").PadRight(maxLength) +
                                              LanguageConsts.EmailFollowUpCompletedText + Environment.NewLine;

                        // Task completed date
                        var completedDate = message.Task.CompleteTime;
                        if (completedDate != null)
                            emailHeader += (LanguageConsts.TaskDateCompleted + ":").PadRight(maxLength) +
                                           ((DateTime) completedDate).ToString(LanguageConsts.DataFormat,
                                               new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine;
                    }
                    else
                    {
                        // Task startdate
                        var startDate = message.Task.StartDate;
                        if (startDate != null)
                            emailHeader += (LanguageConsts.TaskStartDateLabel + ":").PadRight(maxLength) +
                                           ((DateTime) startDate).ToString(LanguageConsts.DataFormat,
                                               new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine;

                        // Task duedate
                        var dueDate = message.Task.DueDate;
                        if (dueDate != null)
                            emailHeader += (LanguageConsts.TaskDueDateLabel + ":").PadRight(maxLength) +
                                           ((DateTime) dueDate).ToString(LanguageConsts.DataFormat,
                                               new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine;

                    }

                    // Empty line
                    emailHeader += Environment.NewLine;
                }

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    emailHeader += (LanguageConsts.EmailCategoriesLabel + ":").PadRight(maxLength) +
                                          String.Join("; ", categories) + Environment.NewLine + Environment.NewLine;
                }

                body = emailHeader + body;
                #endregion
            }

            // Write the body to a file
            File.WriteAllText(eMailFileName, body, Encoding.UTF8);

            return result;
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
            var result = new List<string>();

            // We first always check if there is a RTF body because appointments NEVER have HTML bodies
            var body = message.BodyRtf;
            var htmlBody = false;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                // The RtfToHtmlConverter doesn't support the RTF \objattph tag. So we need to 
                // replace the tag with some text that does survive the conversion. Later on we 
                // will replace these tags with the correct inline image tags
                body = body.Replace("\\objattph", "[OLEATTACHMENT]");
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
                htmlBody = true;
            }

            // When there is no RTF body we try to get the text body
            if (string.IsNullOrEmpty(body))
            {
                body = message.BodyText;
                // When there is no body at all we just make an empty html document
                if (body == null)
                {
                    body = "<html><head></head><body></body></html>";
                    htmlBody = true;
                }
            }

            // Determine the name for the appointment body
            var appointmentFileName = outputFolder +
                                      (!string.IsNullOrEmpty(message.Subject)
                                          ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                          : "appointment") + (htmlBody ? ".htm" : ".txt");

            result.Add(appointmentFileName);

            #region Attachments
            var attachmentList = new List<string>();
            var inlineAttachments = new SortedDictionary<int, string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;

                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attach.FileName));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (htmlBody && attach.RenderingPosition != -1 && IsImageFile(fileInfo.FullName))
                    {
                        inlineAttachments.Add(attach.RenderingPosition, fileInfo.FullName);
                        continue;
                    }

                    inlineAttachments.Add(attach.RenderingPosition, string.Empty);
                    result.Add(fileInfo.FullName);
                }
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;

                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + msg.FileName) + ".msg");
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (msg.RenderingPosition != -1)
                        inlineAttachments.Add(msg.RenderingPosition, string.Empty);
                }

                if (fileInfo == null) continue;

                if (htmlBody)
                    attachmentList.Add("<a href=\"" + HttpUtility.HtmlEncode(fileInfo.Name) + "\">" +
                                       HttpUtility.HtmlEncode(fileInfo.Name) + "</a> (" +
                                       FileManager.GetFileSizeString(fileInfo.Length) + ")");
                else
                    attachmentList.Add(fileInfo.Name + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
            }

            if (htmlBody && hyperlinks)
                foreach (var inlineAttachment in inlineAttachments)
                    body = ReplaceFirstOccurence(body, "[OLEATTACHMENT]", "<img alt=\"\" src=\"" + inlineAttachment.Value + "\">");
            #endregion

            string appointmentHeader;

            if (htmlBody)
            {
                #region Html body
                // Start of table
                appointmentHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // Subject
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentSubject + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Location
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentLocation + ":</td><td>" + message.Appointment.Location + "</td></tr>" + Environment.NewLine;

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Start
                if (message.Appointment.Start != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentStartDate + ":</td><td>" +
                        ((DateTime)message.Appointment.Start).ToString(LanguageConsts.DataFormat,
                            new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" + Environment.NewLine;
                }

                // End
                if (message.Appointment.End != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentEndDate + ":</td><td>" +
                        ((DateTime)message.Appointment.End).ToString(LanguageConsts.DataFormat,
                            new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                        Environment.NewLine;
                }

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Recurrence type
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentRecurrenceTypeLabel + ":</td><td>" +
                    message.Appointment.RecurrenceTypeText + "</td></tr>" + Environment.NewLine;

                // Recurrence patern
                var recurrencePatern = message.Appointment.RecurrencePatern;
                if (!string.IsNullOrEmpty(recurrencePatern))
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentRecurrencePaternLabel + ":</td><td>" +
                        message.Appointment.RecurrencePatern + "</td></tr>" + Environment.NewLine;
                }

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Status
                var status = message.Appointment.StatusText;
                if (status != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentStatusLabel + ":</td><td>" + status + "</td></tr>" + Environment.NewLine;
                }

                // Appointment organizer (FROM)
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentOrganizerLabel + ":</td><td>" + GetEmailSender(message, hyperlinks, true) +
                    "</td></tr>" + Environment.NewLine;

                // Mandatory participants (TO)
                appointmentHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentMandatoryParticipantsLabel + ":</td><td>" +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, true) + "</td></tr>" +
                    Environment.NewLine;

                // Optional participants (CC)
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentOptionalParticipantsLabel + ":</td><td>" + cc + "</td></tr>" + Environment.NewLine;

                // Empty line
                appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories) + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.ImportanceLabel + ":</td><td>" + importance + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    appointmentHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentAttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList) +
                        "</td></tr>" + Environment.NewLine;

                    // Empty line
                    appointmentHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // End of table + empty line
                appointmentHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, appointmentHeader);
                #endregion
            }
            else
            {
                #region Text body
                // Read all the language consts and get the longest string
                var languageConsts = new List<string>
                {
                    LanguageConsts.AppointmentSubject,
                    LanguageConsts.AppointmentLocation,
                    LanguageConsts.AppointmentStartDate,
                    LanguageConsts.AppointmentEndDate,
                    LanguageConsts.AppointmentRecurrenceTypeLabel,
                    LanguageConsts.AppointmentStatusLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentRecurrencePaternLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentMandatoryParticipantsLabel,
                    LanguageConsts.AppointmentOptionalParticipantsLabel,
                    LanguageConsts.AppointmentCategoriesLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                };

                var maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max();

                // Subject
                appointmentHeader = (LanguageConsts.AppointmentSubject + ":").PadRight(maxLength) + message.Subject + Environment.NewLine;

                // Location + empty line
                appointmentHeader += (LanguageConsts.AppointmentLocation + ":").PadRight(maxLength) +
                                     message.Appointment.Location + Environment.NewLine + Environment.NewLine;

                // Start
                if (message.Appointment.Start != null)
                {
                    appointmentHeader += (LanguageConsts.AppointmentStartDate + ":").PadRight(maxLength) +
                                         ((DateTime)message.Appointment.Start).ToString(LanguageConsts.DataFormat,
                                             new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine;
                }

                // End + empty line
                if (message.Appointment.End != null)
                {
                    appointmentHeader += (LanguageConsts.AppointmentEndDate + ":").PadRight(maxLength) +
                                         ((DateTime)message.Appointment.End).ToString(LanguageConsts.DataFormat,
                                             new CultureInfo(LanguageConsts.DateFormatCulture)) + Environment.NewLine +
                                         Environment.NewLine;
                }

                // Recurrence type
                appointmentHeader += (LanguageConsts.AppointmentRecurrenceTypeLabel + ":").PadRight(maxLength) +
                                     message.Appointment.RecurrenceTypeText + Environment.NewLine;

                // Recurrence patern
                var recurrencePatern = message.Appointment.RecurrencePatern;
                if (!string.IsNullOrEmpty(recurrencePatern))
                {
                    appointmentHeader += (LanguageConsts.AppointmentRecurrencePaternLabel + ":").PadRight(maxLength) +
                                     message.Appointment.RecurrencePatern + Environment.NewLine;
                }

                // Empty line
                appointmentHeader += Environment.NewLine;

                // Status
                var status = message.Appointment.StatusText;
                if (status != null)
                {
                    appointmentHeader += (LanguageConsts.AppointmentStatusLabel + ":").PadRight(maxLength) +
                                         status + Environment.NewLine;
                }

                // Appointment organizer (FROM)
                appointmentHeader += (LanguageConsts.AppointmentOrganizerLabel + ":").PadRight(maxLength) +
                     GetEmailSender(message, hyperlinks, false) + Environment.NewLine;

                // Mandatory participants (TO)
                appointmentHeader += (LanguageConsts.AppointmentMandatoryParticipantsLabel + ":").PadRight(maxLength) +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, false) + Environment.NewLine;

                // Optional participants (CC)
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    appointmentHeader +=
                        (LanguageConsts.AppointmentOptionalParticipantsLabel + ":").PadRight(maxLength) + cc +
                        Environment.NewLine;

                // Empty line
                appointmentHeader += Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    appointmentHeader +=
                        (LanguageConsts.AppointmentCategoriesLabel + ":").PadRight(maxLength) + String.Join("; ", categories) +
                        Environment.NewLine + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    appointmentHeader +=
                        (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) + importance + Environment.NewLine +
                        Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    appointmentHeader +=
                        (LanguageConsts.AppointmentAttachmentsLabel + ":").PadRight(maxLength) +
                        string.Join(", ", attachmentList) + Environment.NewLine;
                }

                appointmentHeader += Environment.NewLine;

                body = appointmentHeader + body;
                #endregion
            }

            // Write the body to a file
            File.WriteAllText(appointmentFileName, body, Encoding.UTF8);

            return result;
        }
        #endregion

        #region WriteTask
        /// <summary>
        /// Writes the task body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteTask(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            throw new NotImplementedException("Todo write task code");
            // TODO: Rewrite this code so that an correct task is written

            var result = new List<string>();

            // We first always check if there is a RTF body because appointments NEVER have HTML bodies
            var body = message.BodyRtf;
            var htmlBody = false;
            
            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
            }

            // Determine the name for the task body
            // Determine the name for the appointment body
            var taskFileName = outputFolder +
                                      (!string.IsNullOrEmpty(message.Subject)
                                          ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                          : "task") + (htmlBody ? ".htm" : ".txt");
            result.Add(taskFileName);

            #region Attachments
            var attachmentList = new List<string>();
            var inlineAttachments = new SortedDictionary<int, string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;

                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attach.FileName));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (htmlBody && attach.RenderingPosition != -1 && IsImageFile(fileInfo.FullName))
                    {
                        inlineAttachments.Add(attach.RenderingPosition, fileInfo.FullName);
                        continue;
                    }

                    inlineAttachments.Add(attach.RenderingPosition, string.Empty);
                    result.Add(fileInfo.FullName);
                }
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;

                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + msg.FileName) + ".msg");
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (msg.RenderingPosition != -1)
                        inlineAttachments.Add(msg.RenderingPosition, string.Empty);
                }

                if (fileInfo == null) continue;

                if (htmlBody)
                    attachmentList.Add("<a href=\"" + HttpUtility.HtmlEncode(fileInfo.Name) + "\">" +
                                       HttpUtility.HtmlEncode(fileInfo.Name) + "</a> (" +
                                       FileManager.GetFileSizeString(fileInfo.Length) + ")");
                else
                    attachmentList.Add(fileInfo.Name + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
            }

            if (htmlBody && hyperlinks)
                foreach (var inlineAttachment in inlineAttachments)
                    body = ReplaceFirstOccurence(body, "[OLEATTACHMENT]", "<img alt=\"\" src=\"" + inlineAttachment.Value + "\">");
            #endregion

            string taskHeader;

            if (htmlBody)
            {
                #region Html body
                // Start of table
                taskHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // Subject
                taskHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.TaskSubject + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;


                // When complete
                if (message.Task.Complete != null && (bool) message.Task.Complete)
                {
                    taskHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailFollowUpStatusLabel + ":</td><td>" +
                        LanguageConsts.EmailFollowUpCompletedText +
                        "</td></tr>" + Environment.NewLine;

                    // Task completed date
                    var completedDate = message.Task.CompleteTime;
                    if (completedDate != null)
                        taskHeader +=
                            "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                            LanguageConsts.TaskDateCompleted + ":</td><td>" +
                            ((DateTime) completedDate).ToString(LanguageConsts.DataFormat,
                                new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                            Environment.NewLine;
                }
                else
                {
                    // Task startdate
                    var startDate = message.Task.StartDate;
                    if (startDate != null)
                        taskHeader +=
                            "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                            LanguageConsts.TaskStartDateLabel + ":</td><td>" +
                            ((DateTime) startDate).ToString(LanguageConsts.DataFormat,
                                new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                            Environment.NewLine;

                    // Task duedate
                    var dueDate = message.Task.DueDate;
                    if (dueDate != null)
                        taskHeader +=
                            "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                            LanguageConsts.TaskDueDateLabel + ":</td><td>" +
                            ((DateTime) dueDate).ToString(LanguageConsts.DataFormat,
                                new CultureInfo(LanguageConsts.DateFormatCulture)) + "</td></tr>" +
                            Environment.NewLine;

                }

                // Empty line
                taskHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    taskHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories) + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    taskHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    taskHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.ImportanceLabel + ":</td><td>" + importance + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    taskHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    taskHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentAttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList) +
                        "</td></tr>" + Environment.NewLine;

                    // Empty line
                    taskHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // End of table + empty line
                taskHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, taskHeader);
                #endregion
            }
            else
            {
                #region Text body
                // Read all the language consts and get the longest string
                var languageConsts = new List<string>
                {
                    LanguageConsts.AppointmentSubject,
                    LanguageConsts.AppointmentLocation,
                    LanguageConsts.AppointmentStartDate,
                    LanguageConsts.AppointmentEndDate,
                    LanguageConsts.AppointmentRecurrenceTypeLabel,
                    LanguageConsts.AppointmentStatusLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentRecurrencePaternLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentMandatoryParticipantsLabel,
                    LanguageConsts.AppointmentOptionalParticipantsLabel,
                    LanguageConsts.AppointmentCategoriesLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                };

                var maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max();

                // Subject
                taskHeader = (LanguageConsts.AppointmentSubject + ":").PadRight(maxLength) + message.Subject + Environment.NewLine;

                // Status
                var status = message.Appointment.StatusText;
                if (status != null)
                {
                    taskHeader += (LanguageConsts.AppointmentStatusLabel + ":").PadRight(maxLength) +
                                         status + Environment.NewLine;
                }

                // Appointment organizer (FROM)
                taskHeader += (LanguageConsts.AppointmentOrganizerLabel + ":").PadRight(maxLength) +
                     GetEmailSender(message, hyperlinks, false) + Environment.NewLine;

                // Mandatory participants (TO)
                taskHeader += (LanguageConsts.AppointmentMandatoryParticipantsLabel + ":").PadRight(maxLength) +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, false) + Environment.NewLine;

                // Optional participants (CC)
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    taskHeader +=
                        (LanguageConsts.AppointmentOptionalParticipantsLabel + ":").PadRight(maxLength) + cc +
                        Environment.NewLine;

                // Empty line
                taskHeader += Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    taskHeader +=
                        (LanguageConsts.AppointmentCategoriesLabel + ":").PadRight(maxLength) + String.Join("; ", categories) +
                        Environment.NewLine + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    taskHeader +=
                        (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) + importance + Environment.NewLine +
                        Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    taskHeader +=
                        (LanguageConsts.AppointmentAttachmentsLabel + ":").PadRight(maxLength) +
                        string.Join(", ", attachmentList) + Environment.NewLine;
                }

                taskHeader += Environment.NewLine;

                body = taskHeader + body;
                #endregion
            }

            // Write the body to a file
            File.WriteAllText(taskFileName, body, Encoding.UTF8);

            return result;
        }
        #endregion

        #region WriteContact
        /// <summary>
        /// Writes the task body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private List<string> WriteContact(Storage.Message message, string outputFolder, bool hyperlinks)
        {
            throw new NotImplementedException("Todo write contact code");
            // TODO: Rewrite this code so that an correct task is written

            var result = new List<string>();

            // We first always check if there is a RTF body because appointments NEVER have HTML bodies
            var body = message.BodyRtf;
            var htmlBody = false;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                // The RtfToHtmlConverter doesn't support the RTF \objattph tag. So we need to 
                // replace the tag with some text that does survive the conversion. Later on we 
                // will replace these tags with the correct inline image tags
                body = body.Replace("\\objattph", "[OLEATTACHMENT]");
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
                htmlBody = true;
            }

            // When there is no RTF body we try to get the text body
            if (string.IsNullOrEmpty(body))
            {
                body = message.BodyText;
                // When there is no body at all we just make an empty html document
                if (body == null)
                {
                    body = "<html><head></head><body></body></html>";
                    htmlBody = true;
                }
            }

            // Determine the name for the task body
            // Determine the name for the appointment body
            var taskFileName = outputFolder +
                                      (!string.IsNullOrEmpty(message.Subject)
                                          ? FileManager.RemoveInvalidFileNameChars(message.Subject)
                                          : "contact") + (htmlBody ? ".htm" : ".txt");

            result.Add(taskFileName);

            #region Attachments
            var attachmentList = new List<string>();
            var inlineAttachments = new SortedDictionary<int, string>();

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;

                if (attachment is Storage.Attachment)
                {
                    var attach = (Storage.Attachment)attachment;
                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attach.FileName));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (htmlBody && attach.RenderingPosition != -1 && IsImageFile(fileInfo.FullName))
                    {
                        inlineAttachments.Add(attach.RenderingPosition, fileInfo.FullName);
                        continue;
                    }

                    inlineAttachments.Add(attach.RenderingPosition, string.Empty);
                    result.Add(fileInfo.FullName);
                }
                else if (attachment is Storage.Message)
                {
                    var msg = (Storage.Message)attachment;

                    fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + msg.FileName) + ".msg");
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);

                    // Check if the attachment has a render position. This property is only filled when the
                    // body is RTF and the attachment is made inline
                    if (msg.RenderingPosition != -1)
                        inlineAttachments.Add(msg.RenderingPosition, string.Empty);
                }

                if (fileInfo == null) continue;

                if (htmlBody)
                    attachmentList.Add("<a href=\"" + HttpUtility.HtmlEncode(fileInfo.Name) + "\">" +
                                       HttpUtility.HtmlEncode(fileInfo.Name) + "</a> (" +
                                       FileManager.GetFileSizeString(fileInfo.Length) + ")");
                else
                    attachmentList.Add(fileInfo.Name + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
            }

            if (htmlBody && hyperlinks)
                foreach (var inlineAttachment in inlineAttachments)
                    body = ReplaceFirstOccurence(body, "[OLEATTACHMENT]", "<img alt=\"\" src=\"" + inlineAttachment.Value + "\">");
            #endregion

            string contactHeader;

            if (htmlBody)
            {
                #region Html body
                // Start of table
                contactHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                // Subject
                contactHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentSubject + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;


                // Empty line
                contactHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Appointment organizer (FROM)
                contactHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                    LanguageConsts.AppointmentOrganizerLabel + ":</td><td>" + GetEmailSender(message, hyperlinks, true) +
                    "</td></tr>" + Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    contactHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.EmailCategoriesLabel + ":</td><td>" + String.Join("; ", categories) + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    contactHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    contactHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.ImportanceLabel + ":</td><td>" + importance + "</td></tr>" + Environment.NewLine;

                    // Empty line
                    contactHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    contactHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.AppointmentAttachmentsLabel + ":</td><td>" + string.Join(", ", attachmentList) +
                        "</td></tr>" + Environment.NewLine;

                    // Empty line
                    contactHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                }

                // End of table + empty line
                contactHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, contactHeader);
                #endregion
            }
            else
            {
                #region Text body
                // Read all the language consts and get the longest string
                var languageConsts = new List<string>
                {
                    LanguageConsts.AppointmentSubject,
                    LanguageConsts.AppointmentLocation,
                    LanguageConsts.AppointmentStartDate,
                    LanguageConsts.AppointmentEndDate,
                    LanguageConsts.AppointmentRecurrenceTypeLabel,
                    LanguageConsts.AppointmentStatusLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentRecurrencePaternLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentMandatoryParticipantsLabel,
                    LanguageConsts.AppointmentOptionalParticipantsLabel,
                    LanguageConsts.AppointmentCategoriesLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                };

                var maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max();

                // Subject
                contactHeader = (LanguageConsts.AppointmentSubject + ":").PadRight(maxLength) + message.Subject + Environment.NewLine;

                // Status
                var status = message.Appointment.StatusText;
                if (status != null)
                {
                    contactHeader += (LanguageConsts.AppointmentStatusLabel + ":").PadRight(maxLength) +
                                         status + Environment.NewLine;
                }

                // Appointment organizer (FROM)
                contactHeader += (LanguageConsts.AppointmentOrganizerLabel + ":").PadRight(maxLength) +
                     GetEmailSender(message, hyperlinks, false) + Environment.NewLine;

                // Mandatory participants (TO)
                contactHeader += (LanguageConsts.AppointmentMandatoryParticipantsLabel + ":").PadRight(maxLength) +
                    GetEmailRecipients(message, Storage.Recipient.RecipientType.To, hyperlinks, false) + Environment.NewLine;

                // Optional participants (CC)
                var cc = GetEmailRecipients(message, Storage.Recipient.RecipientType.Cc, hyperlinks, false);
                if (cc != string.Empty)
                    contactHeader +=
                        (LanguageConsts.AppointmentOptionalParticipantsLabel + ":").PadRight(maxLength) + cc +
                        Environment.NewLine;

                // Empty line
                contactHeader += Environment.NewLine;

                // Categories
                var categories = message.Categories;
                if (categories != null)
                {
                    contactHeader +=
                        (LanguageConsts.AppointmentCategoriesLabel + ":").PadRight(maxLength) + String.Join("; ", categories) +
                        Environment.NewLine + Environment.NewLine;
                }

                // Urgent
                var importance = message.Importance;
                if (importance != null)
                {
                    contactHeader +=
                        (LanguageConsts.ImportanceLabel + ":").PadRight(maxLength) + importance + Environment.NewLine +
                        Environment.NewLine;
                }

                // Attachments
                if (attachmentList.Count != 0)
                {
                    contactHeader +=
                        (LanguageConsts.AppointmentAttachmentsLabel + ":").PadRight(maxLength) +
                        string.Join(", ", attachmentList) + Environment.NewLine;
                }

                contactHeader += Environment.NewLine;

                body = contactHeader + body;
                #endregion
            }

            // Write the body to a file
            File.WriteAllText(taskFileName, body, Encoding.UTF8);

            return result;
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
            var result = new List<string>();
            string stickyNoteFile;
            var stickyNoteHeader = string.Empty;

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

                stickyNoteHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine;

                if (message.SentOn != null)
                    stickyNoteHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +
                        LanguageConsts.StickyNoteDateLabel + ":</td><td>" +
                        ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormat) + "</td></tr>" +
                        Environment.NewLine;

                // Empty line
                stickyNoteHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;
                
                // End of table + empty line
                stickyNoteHeader += "</table><br/>" + Environment.NewLine;

                body = InjectHeader(body, stickyNoteHeader);
            }
            else
            {
                body = message.BodyText ?? string.Empty;

                // Sent on
                if (message.SentOn != null)
                    stickyNoteHeader +=
                        (LanguageConsts.StickyNoteDateLabel + ":") + ((DateTime) message.SentOn).ToString(LanguageConsts.DataFormat) + Environment.NewLine;

                body = stickyNoteHeader + body;
                stickyNoteFile = outputFolder + (!string.IsNullOrEmpty(message.Subject) ? FileManager.RemoveInvalidFileNameChars(message.Subject) : "stickynote") + ".txt";   
            }

            // Write the body to a file
            File.WriteAllText(stickyNoteFile, body, Encoding.UTF8);
            result.Add(stickyNoteFile);
            return result;
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
        /// Change the E-mail sender addresses to a human readable format
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
                    recipients.Add(new Recipient { EmailAddress = recipient.Email, DisplayName = recipient.DisplayName });
            }

            if (recipients.Count == 0 && message.Headers != null)
            {
                switch (type)
                {
                    case Storage.Recipient.RecipientType.To:
                        if (message.Headers.To != null)
                            recipients.AddRange(message.Headers.To.Select(to => new Recipient {EmailAddress = to.Address, DisplayName = to.DisplayName}));
                        break;

                    case Storage.Recipient.RecipientType.Cc:
                        if (message.Headers.Cc != null)
                            recipients.AddRange(message.Headers.Cc.Select(cc => new Recipient { EmailAddress = cc.Address, DisplayName = cc.DisplayName }));
                        break;

                    case Storage.Recipient.RecipientType.Bcc:
                        if (message.Headers.Bcc != null)
                            recipients.AddRange(message.Headers.Bcc.Select(bcc => new Recipient { EmailAddress = bcc.Address, DisplayName = bcc.DisplayName }));
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
                    if(!string.IsNullOrEmpty(displayName))
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

        #region GetInnerException
        /// <summary>
        /// Get the complete inner exception tree
        /// </summary>
        /// <param name="e">The exception object</param>
        /// <returns></returns>
        private static string GetInnerException(Exception e)
        {
            var exception = e.Message + "\n";
            if (e.InnerException != null)
                exception += GetInnerException(e.InnerException);
            return exception;
        }
        #endregion

        #region IsImageFile
        /// <summary>
        /// Returns true when the given fileName is an image
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool IsImageFile(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;

            var extension = Path.GetExtension(fileName);
            if (!string.IsNullOrEmpty(extension))
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".JPG":
                    case ".JPEG":
                    case ".TIF":
                    case ".TIFF":
                    case ".GIF":
                    case ".BMP":
                    case ".PNG":
                        return true;
                }
            }

            return false;
        }
        #endregion
    }
}
