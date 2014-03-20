using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
            var result = new List<string>();
            _errorMessage = string.Empty;

            try
            {
                using (var messageStream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                {
                    // Read MSG file from a stream
                    using (var message = new Storage.Message(messageStream))
                    {
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
                            var fileName = string.Empty;
                            if (attachment.GetType() == typeof (Storage.Attachment))
                            {
                                var attach = (Storage.Attachment) attachment;
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
                            else if (attachment.GetType() == typeof (Storage.Message))
                            {
                                var msg = (Storage.Message) attachment;
                                fileName =
                                    FileManager.FileExistsMakeNew(outputFolder +
                                                                  FileManager.RemoveInvalidFileNameChars(msg.Subject) +
                                                                  ".msg");
                                result.Add(fileName);
                                msg.Save(fileName);
                            }

                            if (attachments == string.Empty)
                                attachments = Path.GetFileName(fileName);
                            else
                                attachments += ", " + Path.GetFileName(fileName);
                        }

                        string outlookHeader;

                        if (htmlBody)
                        {
                            // Add an outlook style header into the HTML body.
                            // Change this code to the language you need. 
                            // Currently it is written in ENGLISH
                            outlookHeader =
                                "<TABLE cellSpacing=0 cellPadding=0 width=\"100%\" border=0 style=\"font-family: 'Times New Roman'; font-size: 12pt;\"\\>" +
                                Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>From:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" +
                                GetEmailSender(message) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>To:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" +
                                GetEmailRecipients(message, Storage.RecipientType.To) + "</TD></TR>" +
                                Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Sent on:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" +
                                (message.SentOn != null
                                    ? ((DateTime) message.SentOn).ToString("dd-MM-yyyy HH:mm:ss")
                                    : string.Empty) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Received on:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" +
                                (message.ReceivedOn != null
                                    ? ((DateTime) message.ReceivedOn).ToString("dd-MM-yyyy HH:mm:ss")
                                    : string.Empty) + "</TD></TR>";

                            // CC
                            var cc = GetEmailRecipients(message, Storage.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader +=
                                    "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>CC:</STRONG></TD><TD style=\"height: 18px\">" +
                                    cc + "</TD></TR>" + Environment.NewLine;

                            // Subject
                            outlookHeader +=
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Subject:</STRONG></TD><TD style=\"height: 18px\">" +
                                message.Subject + "</TD></TR>" + Environment.NewLine;

                            // Empty line
                            outlookHeader += "<TR><TD colspan=\"2\" style=\"height: 18px\">&nbsp</TD></TR>" +
                                             Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader +=
                                    "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Attachments:</STRONG></TD><TD style=\"height: 18px\">" +
                                    attachments + "</TD></TR>" + Environment.NewLine;

                            //  End of table + empty line
                            outlookHeader += "</TABLE><BR>" + Environment.NewLine;

                            body = InjectOutlookHeader(body, outlookHeader);
                        }
                        else
                        {
                            // Add an outlook style header into the Text body. 
                            // Change this code to the language you need. 
                            // Currently it is written in ENGLISH
                            outlookHeader =
                                "From:\t\t" + GetEmailSender(message) + Environment.NewLine +
                                "To:\t\t" + GetEmailRecipients(message, Storage.RecipientType.To) + Environment.NewLine +
                                "Sent on:\t" +
                                (message.SentOn != null
                                    ? ((DateTime) message.SentOn).ToString("dd-MM-yyyy HH:mm:ss")
                                    : string.Empty) + Environment.NewLine +
                                "Received on:\t" +
                                (message.ReceivedOn != null
                                    ? ((DateTime) message.ReceivedOn).ToString("dd-MM-yyyy HH:mm:ss")
                                    : string.Empty) + Environment.NewLine;

                            // CC
                            var cc = GetEmailRecipients(message, Storage.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader += "CC:\t\t" + cc + Environment.NewLine;

                            outlookHeader += "Subject:\t" + message.Subject + Environment.NewLine + Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader += "Attachments:\t" + attachments + Environment.NewLine +
                                                 Environment.NewLine;

                            body = outlookHeader + body;
                        }

                        // Write the body to a file
                        File.WriteAllText(eMailFileName, body, Encoding.UTF8);
                    }
                }
            }
            catch (Exception e)
            {
                _errorMessage = GetInnerException(e);
                return new string[0];
            }

            return result.ToArray();
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

        #region GetEmailSender
        /// <summary>
        /// Change the E-mail sender addresses to a human readable format
        /// </summary>
        /// <param name="message">The Storage.Message object</param>
        /// <param name="convertToHref">When true the E-mail addresses are converted to hyperlinks</param>
        /// <returns></returns>
        private static string GetEmailSender(Storage.Message message, bool convertToHref = false)
        {
            var output = string.Empty;

            if (message == null) return string.Empty;

            var eMail = message.Sender.Email;
            if (string.IsNullOrEmpty(eMail))
            {
                if (message.Headers != null && message.Headers.From != null)
                    eMail = message.Headers.From.Address;
            }

            var displayName = message.Sender.DisplayName;
            if (string.IsNullOrEmpty(displayName))
            {
                if (message.Headers != null && message.Headers.From != null)
                    displayName = message.Headers.From.DisplayName;
            }

            if (string.IsNullOrEmpty(eMail))
                convertToHref = false;

            if (convertToHref)
                output += "<a href=\"mailto:" + eMail + "\">" +
                          (!string.IsNullOrEmpty(displayName)
                              ? displayName
                              : eMail) + "</a>";

            else
            {
                if (string.IsNullOrEmpty(eMail))
                {
                    output += !string.IsNullOrEmpty(displayName)
                        ? displayName
                        : string.Empty;
                }
                else
                {
                    output += eMail +
                              (!string.IsNullOrEmpty(displayName)
                                  ? " (" + displayName + ")"
                                  : string.Empty);
                }
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
                                                 bool convertToHref = false)
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

                var convert = convertToHref;

                if (convert && string.IsNullOrEmpty(recipient.EmailAddress))
                    convert = false;

                if (convert)
                {
                    output += "<a href=\"mailto:" + message.Sender.Email + "\">" +
                              (!string.IsNullOrEmpty(message.Sender.DisplayName)
                                  ? recipient.DisplayName
                                  : recipient.EmailAddress) + "</a>";
                }
                else
                {
                    if (string.IsNullOrEmpty(recipient.EmailAddress))
                    {
                        output += !string.IsNullOrEmpty(recipient.DisplayName)
                            ? recipient.DisplayName
                            : string.Empty;
                    }
                    else
                    {
                        output += recipient.EmailAddress +
                                  (!string.IsNullOrEmpty(recipient.DisplayName)
                                      ? " (" + recipient.DisplayName + ")"
                                      : string.Empty);
                    }                    
                }
            }

            return output;
        }
        #endregion

        #region InjectOutlookHeader
        /// <summary>
        /// Inject an outlook style header into the email body
        /// </summary>
        /// <param name="eMail"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        private string InjectOutlookHeader(string eMail, string header)
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
