using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DocumentServices.Modules.Readers.MsgReader.Outlook;

namespace DocumentServices.Modules.Readers.MsgReader
{
    public interface IReader
    {
        /// <summary>
        /// Extract the input msg file to the given output folder
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        /// <param name="outputFolder">The folder where to extract the msg file</param>
        /// <returns>String array containing the message body and its (inline) attachments</returns>
        [DispId(1)]
        string[] ExtractToFolder(string inputFile, string outputFolder);

        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        string GetErrorMessage();
    }

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

        #region internal class
        /// <summary>
        /// Used as a placeholder for the recipients from the MSG file itself or from the "internet"
        /// headers when this message is send outside an Exchange system
        /// </summary>
        internal class Recipient
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
        /// <returns>String array containing the message body and its (inline) attachments</returns>
        public string[] ExtractToFolder(string inputFile, string outputFolder)
        {
            outputFolder = FileManager.CheckForBackSlash(outputFolder);
            _errorMessage = string.Empty;

            try
            {
                using (var messageStream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                {
                    using (var message = new Storage.Message(messageStream))
                    {
                        switch (message.Type)
                        {
                            case "IPM.Note":
                                return WriteEmail(message, outputFolder).ToArray();

                            case "IPM.Appointment":
                                return WriteAppointment(message, outputFolder).ToArray();
                        }
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

        #region WriteEmail
        /// <summary>
        /// Writes the body of the MSG E-mail to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <returns></returns>
        private List<string> WriteEmail(Storage.Message message, string outputFolder)
        {
            var result = new List<string>();

            // Read MSG file from a stream
            // We first always check if there is a HTML body
            var body = message.BodyHtml;
            var htmlBody = true;

            // Determine the name for the E-mail body
            var eMailFileName = outputFolder + "email" + (body != null ? ".htm" : ".txt");
            result.Add(eMailFileName);

            if (body == null)
            {
                // When there is not HTML body found then try to get the text body
                body = message.BodyText;
                htmlBody = false;
            }

            var attachments = string.Empty;

            foreach (var attachment in message.Attachments)
            {
                FileInfo fileInfo = null;
                
                if (attachment.GetType() == typeof (Storage.Attachment))
                {
                    var attach = (Storage.Attachment) attachment;
                    fileInfo = new FileInfo(
                        FileManager.FileExistsMakeNew(outputFolder +
                                                      FileManager.RemoveInvalidFileNameChars(attach.Filename)));
                    File.WriteAllBytes(fileInfo.FullName, attach.Data);

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) &&
                        body.Contains(attach.ContentId))
                    {
                        body = body.Replace("cid:" + attach.ContentId, fileInfo.FullName);
                        continue;
                    }

                    result.Add(fileInfo.FullName);
                }
                else if (attachment.GetType() == typeof (Storage.Message))
                {
                    var msg = (Storage.Message) attachment;
                    fileInfo = new FileInfo(
                        FileManager.FileExistsMakeNew(outputFolder +
                                                      FileManager.RemoveInvalidFileNameChars(msg.Subject) + ".msg"));
                    result.Add(fileInfo.FullName);
                    msg.Save(fileInfo.FullName);
                }

                if (fileInfo == null) continue;

                if (string.IsNullOrEmpty(attachments))
                    attachments = fileInfo.Name;
                else
                    attachments += ", " + fileInfo.Name;
            }

            string outlookEmailHeader;

            // The labels used in the header of the E-mail, change these to your own language when needed
            const string fromLabel = "From";
            const string sentOnLabel = "Sent on";
            const string toLabel = "To";
            const string receivedOnLabel = "Received on";
            const string subjectLabel = "Subject";
            const string ccLabel = "CC";
            const string attachmentsLabel = "Attachments";
            const string categoriesLabel = "Categories";

            // The date format used in the header of the E-mail
            const string dataFormat = "dd-MM-yyyy HH:mm:ss";

            // When true then "to" and "from" E-mail addresses are converted to "mailto:" hyperlinks when there is an html body
            const bool convertEmailsToHyperLinks = false;
            
            if (htmlBody)
            {
                // Add an outlook style header into the HTML body.
                outlookEmailHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine +
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + fromLabel + ":</td><td>" + GetEmailSender(message, convertEmailsToHyperLinks) + "</td></tr>" + Environment.NewLine;

                if (message.SentOn != null)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + sentOnLabel + ":</td><td>" + ((DateTime)message.SentOn).ToString(dataFormat) + "</td></tr>" + Environment.NewLine;

                outlookEmailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + toLabel + ":</td><td>" + GetEmailRecipients(message, Storage.RecipientType.To, convertEmailsToHyperLinks) + "</td></tr>" + Environment.NewLine;

                if (message.ReceivedOn != null)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + receivedOnLabel + ":</td><td>" + ((DateTime)message.ReceivedOn).ToString(dataFormat) + "</td></tr>" + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.RecipientType.Cc, convertEmailsToHyperLinks);
                if (cc != string.Empty)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" +ccLabel + ":</td><td>" + cc + "</td></tr>" + Environment.NewLine;

                // Subject
                outlookEmailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + subjectLabel + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Empty line
                outlookEmailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Attachments
                if (attachments != string.Empty)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + attachmentsLabel + ":</td><td>" + attachments + "</td></tr>" + Environment.NewLine;

                var categories = message.Categories;
                if (categories != null)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + categoriesLabel + ":</td><td>" + String.Join("; ", categories) + "</td></tr>" + Environment.NewLine;


                // End of table + empty line
                outlookEmailHeader += "</table><br/>" + Environment.NewLine;

                body = InjectOutlookEmailHeader(body, outlookEmailHeader);
            }
            else
            {
                outlookEmailHeader =
                    fromLabel + ":\t\t" + GetEmailSender(message, false) + Environment.NewLine;
                    
                if (message.SentOn != null)
                    outlookEmailHeader +=
                        sentOnLabel + ":\t" + ((DateTime)message.SentOn).ToString(dataFormat) + Environment.NewLine;

                outlookEmailHeader +=
                    toLabel + ":\t\t" + GetEmailRecipients(message, Storage.RecipientType.To, false) + Environment.NewLine;

                if (message.ReceivedOn != null)
                    outlookEmailHeader +=
                        receivedOnLabel + ":\t" + ((DateTime)message.ReceivedOn).ToString(dataFormat) + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.RecipientType.Cc, false);
                if (cc != string.Empty)
                    outlookEmailHeader += ccLabel + ":\t\t" + cc + Environment.NewLine;

                outlookEmailHeader += subjectLabel + ":\t" + message.Subject + Environment.NewLine + Environment.NewLine;

                // Attachments
                if (attachments != string.Empty)
                    outlookEmailHeader += attachmentsLabel + ":\t" + attachments + Environment.NewLine + Environment.NewLine;

                body = outlookEmailHeader + body;
            }

            // Write the body to a file
            File.WriteAllText(eMailFileName, body, Encoding.UTF8);

            return result;
        }
        #endregion

        #region WriteAppointment
        /// <summary>
        /// Writes the body of the MSG Appointment to html or text and extracts all the attachments. The
        /// result is return as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <returns></returns>
        private List<string> WriteAppointment(Storage.Message message, string outputFolder)
        {
            // TODO: Rewrite this code so that an correct appointment is written

            var result = new List<string>();

            // Read MSG file from a stream
            // We first always check if there is a RTF body because appointments never have HTML bodies
            var body = message.BodyRtf;

            // If the body is not null then we convert it to HTML
            if (body != null)
            {
                var converter = new RtfToHtmlConverter();
                body = converter.ConvertRtfToHtml(body);
            }

            var htmlBody = true;

            // Determine the name for the appointment body
            var appointmentFileName = outputFolder + "appointment" + (body != null ? ".htm" : ".txt");
            result.Add(appointmentFileName);

            if (body == null)
            {
                // When there is not RTF body found then try to get the text body
                body = message.BodyText;
                htmlBody = false;
            }

            var attachments = string.Empty;

            foreach (var attachment in message.Attachments)
            {
                var fileName = string.Empty;
                if (attachment.GetType() == typeof(Storage.Attachment))
                {
                    var attach = (Storage.Attachment)attachment;
                    fileName =
                        FileManager.FileExistsMakeNew(outputFolder +
                                                      FileManager.RemoveInvalidFileNameChars(attach.Filename));
                    File.WriteAllBytes(fileName, attach.Data);

                    // When we find an inline attachment we have to replace the CID tag inside the html body
                    // with the name of the inline attachment. But before we do this we check if the CID exists.
                    // When the CID does not exists we treat the inline attachment as a normal attachment
                    if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) &&
                        body.Contains(attach.ContentId))
                    {
                        body = body.Replace("cid:" + attach.ContentId, fileName);
                        continue;
                    }

                    result.Add(fileName);
                }
                else if (attachment.GetType() == typeof(Storage.Message))
                {
                    var msg = (Storage.Message)attachment;
                    fileName =
                        FileManager.FileExistsMakeNew(outputFolder +
                                                      FileManager.RemoveInvalidFileNameChars(msg.Subject) + ".msg");
                    result.Add(fileName);
                    msg.Save(fileName);
                }

                if (attachments == string.Empty)
                    attachments = Path.GetFileName(fileName);
                else
                    attachments += ", " + Path.GetFileName(fileName);
            }

            string outlookEmailHeader;

            // The labels used in the header of the E-mail, change these to your own language when needed
            const string fromLabel = "From";
            const string toLabel = "To";
            const string sentOnLabel = "Sent on";
            const string receivedOnLabel = "Received on";
            const string subjectLabel = "Subject";
            const string ccLabel = "CC";
            const string attachmentsLabel = "Attachments";

            // The date format used in the header of the E-mail
            const string dataFormat = "dd-MM-yyyy HH:mm:ss";

            // When true then to E-mails are converted to mailto: hyperlinks when there is an html body
            const bool convertEmailsToHyperLinks = false;
            
            if (htmlBody)
            {
                // Add an outlook style header into the HTML body.
                outlookEmailHeader =
                    "<table style=\"width:100%; font-family: Times New Roman; font-size: 12pt;\">" + Environment.NewLine +
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + fromLabel + ":</td><td>" + GetEmailSender(message, convertEmailsToHyperLinks) + "</td></tr>" + Environment.NewLine +
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + toLabel + ":</td><td>" + GetEmailRecipients(message, Storage.RecipientType.To, convertEmailsToHyperLinks) + "</td></tr>" + Environment.NewLine;

                if (message.SentOn != null)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + sentOnLabel + ":</td><td>" + ((DateTime)message.SentOn).ToString(dataFormat) + "</td></tr>" + Environment.NewLine;

                if (message.ReceivedOn != null)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + receivedOnLabel + ":</td><td>" + ((DateTime)message.ReceivedOn).ToString(dataFormat) + "</td></tr>" + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.RecipientType.Cc, convertEmailsToHyperLinks);
                if (cc != string.Empty)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + ccLabel + ":</td><td>" + cc + "</td></tr>" + Environment.NewLine;

                // Subject
                outlookEmailHeader +=
                    "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + subjectLabel + ":</td><td>" + message.Subject + "</td></tr>" + Environment.NewLine;

                // Empty line
                outlookEmailHeader += "<tr><td colspan=\"2\" style=\"height: 18px; \">&nbsp</td></tr>" + Environment.NewLine;

                // Attachments
                if (attachments != string.Empty)
                    outlookEmailHeader +=
                        "<tr style=\"height: 18px; vertical-align: top; \"><td style=\"width: 100px; font-weight: bold; \">" + attachmentsLabel + ":</td><td>" + attachments + "</td></tr>" + Environment.NewLine;

                // End of table + empty line
                outlookEmailHeader += "</table><br/>" + Environment.NewLine;

                body = InjectOutlookEmailHeader(body, outlookEmailHeader);
            }
            else
            {
                outlookEmailHeader =
                    fromLabel + ":\t\t" + GetEmailSender(message, false) + Environment.NewLine +
                    toLabel + ":\t\t" + GetEmailRecipients(message, Storage.RecipientType.To, false) + Environment.NewLine;

                if (message.SentOn != null)
                    outlookEmailHeader +=
                        sentOnLabel + ":\t" + ((DateTime)message.SentOn).ToString(dataFormat) + Environment.NewLine;

                if (message.ReceivedOn != null)
                    outlookEmailHeader +=
                        receivedOnLabel + ":\t" + ((DateTime)message.ReceivedOn).ToString(dataFormat) + Environment.NewLine;

                // CC
                var cc = GetEmailRecipients(message, Storage.RecipientType.Cc, false);
                if (cc != string.Empty)
                    outlookEmailHeader += ccLabel + ":\t\t" + cc + Environment.NewLine;

                outlookEmailHeader += subjectLabel + ":\t" + message.Subject + Environment.NewLine + Environment.NewLine;

                // Attachments
                if (attachments != string.Empty)
                    outlookEmailHeader += attachmentsLabel + ":\t" + attachments + Environment.NewLine + Environment.NewLine;

                body = outlookEmailHeader + body;
            }

            // Write the body to a file
            File.WriteAllText(appointmentFileName, body, Encoding.UTF8);

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
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <returns></returns>
        private static string GetEmailSender(Storage.Message message, bool convertToHref)
        {
            var output = string.Empty;

            if (message == null) return string.Empty;
            
            var tempEmailAddress = message.Sender.Email;
            var tempDisplayName = message.Sender.DisplayName;

            if (string.IsNullOrEmpty(tempEmailAddress) && message.Headers != null && message.Headers.From != null)
                tempEmailAddress = RemoveSingleQuotes(message.Headers.From.Address);
            
            if (string.IsNullOrEmpty(tempDisplayName) && message.Headers != null && message.Headers.From != null)
                tempDisplayName = message.Headers.From.DisplayName;

            string emailAddress = null;
            string displayName = null;

            // Sometimes the E-mail address and displayname get swapped so check if they are valid
            if (IsEmailAddressValid(tempEmailAddress) && !IsEmailAddressValid(tempDisplayName))
            {
                // Keep them as they are
                emailAddress = tempEmailAddress;
                displayName = tempDisplayName;
            }
            else if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
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
            
            if (convertToHref && !string.IsNullOrEmpty(emailAddress))
                output += "<a href=\"mailto:" + emailAddress + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                              ? HttpUtility.HtmlEncode(displayName)
                              : emailAddress) + "</a>";

            else
            {
                output = emailAddress;
                if (!string.IsNullOrEmpty(displayName))
                    output += " <" + displayName + ">";

                if (output != null)
                    output = HttpUtility.HtmlEncode(output);
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
        /// <param name="type">This types says if we want to get the TO's or CC's</param>
        /// <returns></returns>
        private static string GetEmailRecipients(Storage.Message message,
                                                 Storage.RecipientType type,
                                                 bool convertToHref)
        {
            var output = string.Empty;

            var recipients = new List<Recipient>();

            if (message == null)
                return output;

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
                    case Storage.RecipientType.To:
                        foreach (var to in message.Headers.To)
                            recipients.Add(new Recipient { EmailAddress = to.Address, DisplayName = to.DisplayName });
                        break;
        
                    case Storage.RecipientType.Cc:
                        foreach (var cc in message.Headers.Cc)
                            recipients.Add(new Recipient { EmailAddress = cc.Address, DisplayName = cc.DisplayName });
                        break;
                }
            }

            foreach (var recipient in recipients)
            {
                if (output != string.Empty)
                    output += "; ";

                var tempEmailAddress = RemoveSingleQuotes(recipient.EmailAddress);
                var tempDisplayName = RemoveSingleQuotes(recipient.DisplayName);

                if (string.IsNullOrEmpty(tempEmailAddress) && message.Headers != null && message.Headers.From != null)
                    tempEmailAddress = RemoveSingleQuotes(message.Headers.From.Address);

                if (string.IsNullOrEmpty(tempDisplayName) && message.Headers != null && message.Headers.From != null)
                    tempDisplayName = message.Headers.From.DisplayName;

                string emailAddress = null;
                string displayName = null;

                // Sometimes the E-mail address and displayname get swapped so check if they are valid
                if (IsEmailAddressValid(tempEmailAddress) && !IsEmailAddressValid(tempDisplayName))
                {
                    // Keep them as they are
                    emailAddress = tempEmailAddress;
                    displayName = tempDisplayName;
                }
                else if (!IsEmailAddressValid(tempEmailAddress) && IsEmailAddressValid(tempDisplayName))
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

                if (convertToHref && !string.IsNullOrEmpty(emailAddress))
                    output += "<a href=\"mailto:" + emailAddress + "\">" +
                              (!string.IsNullOrEmpty(displayName)
                                  ? HttpUtility.HtmlEncode(displayName)
                                  : emailAddress) + "</a>";

                else
                {
                    output = emailAddress;
                    if (!string.IsNullOrEmpty(displayName))
                        output += " <" + displayName + ">";

                    if (output != null)
                        output = HttpUtility.HtmlEncode(output);
                }
            }

            return output;
        }
        #endregion

        #region InjectOutlookEmailHeader
        /// <summary>
        /// Inject an outlook style header into the email body
        /// </summary>
        /// <param name="eMail"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        private string InjectOutlookEmailHeader(string eMail, string header)
        {
            var temp = eMail.ToUpper();

            var begin = temp.IndexOf("<BODY", StringComparison.Ordinal);

            if (begin > 0)
            {
                begin = temp.IndexOf(">", begin, StringComparison.Ordinal);
                return eMail.Insert(begin + 1, header);
            }

            return header + eMail;
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
    }
}
