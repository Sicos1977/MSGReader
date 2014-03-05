using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web;
using System.Web.SessionState;
using DocumentServices.Modules.Readers.MsgReader;
using DocumentServices.Modules.Readers.MsgReader.Outlook;

namespace MsgViewerWeb
{
    /// <summary>
    /// Summary description for MsgHandler
    /// </summary>
    public class MsgHandler : IHttpHandler, IRequiresSessionState
    {
        private const string VirtualMessagesDir = @"~/FileSystem";

        public string RootMessagesDir { get; set; }

        public void ProcessRequest(HttpContext context)
        {
            RootMessagesDir = context.Server.MapPath(VirtualMessagesDir);
            var msgFileName = context.Request.QueryString.Get("file");
            var msgFileFullName = Path.Combine(RootMessagesDir, msgFileName);
            var attachmentFileName = context.Request.QueryString.Get("attachment");

            if (!string.IsNullOrEmpty(attachmentFileName))
            {
                ServerAttachment(msgFileName, attachmentFileName, context);
                return;
            }

            ServeMsgFile(msgFileFullName, context, msgFileName);
        }

        private void ServerAttachment(string msgFileName, string attachedfileName, HttpContext context)
        {
            var attachmentFolder = Path.Combine(RootMessagesDir, "messages", context.Session.SessionID, Path.GetFileNameWithoutExtension(msgFileName));
            var attachedfileFullName = Path.Combine(attachmentFolder, attachedfileName);
            var ext = Path.GetExtension(attachedfileFullName);

            if (ext.ToLower() == ".msg")
            {
                ServeMsgFile(attachedfileFullName, context, attachedfileName);
                return;
            }
            if (File.Exists(attachedfileFullName))
            {
                context.Response.ContentType = MimeExtensionHelper.GetMimeType(attachedfileFullName);
                context.Response.TransmitFile(attachedfileFullName);
                return;
            }

            throw new HttpException(404, "File does not exist.");
        }

        private void ServeMsgFile(string fileFullName, HttpContext context, string fileName)
        {
            try
            {
                if (File.Exists(fileFullName))
                {
                    var msgFolder = Path.Combine(RootMessagesDir, "messages", context.Session.SessionID, Path.GetFileNameWithoutExtension(fileName));

                    if (!Directory.Exists(msgFolder))
                    {
                        Directory.CreateDirectory(msgFolder);
                    }

                    var emailReader = MessageReaderFactory.CreateMessageReader(context);

                    var files = emailReader.ExtractToFolder(fileFullName, msgFolder);

                    if (files.Length > 0)
                    {
                        //always extracts it to email.html
                        var email = files[0];

                        context.Response.ContentType = MimeExtensionHelper.GetMimeType(email);
                        context.Response.TransmitFile(email);
                        return;
                    }
                    throw new HttpException(404, "File could not be converted.");
                }
                throw new HttpException(404, "File does not exist.");
            }
            catch (HttpException e)
            {
                context.Response.StatusCode = e.GetHttpCode();
                context.Response.Status = e.GetHtmlErrorMessage() ?? e.Message;
                context.Response.End();
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }

    public static class MimeExtensionHelper
    {
        static object locker = new object();
        static object mimeMapping;
        static MethodInfo getMimeMappingMethodInfo;

        static MimeExtensionHelper()
        {
            Type mimeMappingType = Assembly.GetAssembly(typeof(HttpRuntime)).GetType("System.Web.MimeMapping");
            if (mimeMappingType == null)
                throw new SystemException("Couldn't find MimeMapping type");
            getMimeMappingMethodInfo = mimeMappingType.GetMethod("GetMimeMapping", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);
            if (getMimeMappingMethodInfo == null)
                throw new SystemException("Couldn't find GetMimeMapping method");
            if (getMimeMappingMethodInfo.ReturnType != typeof(string))
                throw new SystemException("GetMimeMapping method has invalid return type");
            if (getMimeMappingMethodInfo.GetParameters().Length != 1 && getMimeMappingMethodInfo.GetParameters()[0].ParameterType != typeof(string))
                throw new SystemException("GetMimeMapping method has invalid parameters");
        }
        public static string GetMimeType(string filename)
        {
            lock (locker)
                return (string)getMimeMappingMethodInfo.Invoke(mimeMapping, new object[] { filename });
        }
    }

    public class WebReader : ReaderBase
    {
        public WebReader(string msgHandlerUrl, string applicationPath, string virtualPath)
        {
            MsgHandlerUrl = msgHandlerUrl;
            ApplicationPath = applicationPath;
            VirtualPath = virtualPath;
        }

        public string MsgHandlerUrl { get; private set; }
        protected string VirtualPath { get; private set; }
        public string ApplicationPath { get; private set; }

        public override string[] ExtractToFolder(string inputFile, string outputFolder)
        {
            outputFolder = FileManager.CheckForSlash(outputFolder);
            _errorMessage = string.Empty;

            try
            {
                using (var messageStream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                {
                    // Read MSG file from a stream
                    using (var message = new Storage.Message(messageStream))
                    {
                        var result = new List<string>();
                        // Determine the name for the E-mail body
                        var eMailFileName = Path.Combine(outputFolder,
                                                         "email" + (message.BodyHtml != null ? ".html" : ".txt"));
                        result.Add(eMailFileName);

                        // We first always check if there is a HTML body
                        var body = message.BodyHtml;
                        var htmlBody = true;
                        if (body == null)
                        {
                            // When not found try to get the text body
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
                                                                  FileManager.RemoveInvalidFileNameChars(attach.Filename)
                                                                             .Replace("&", "[and]"));
                                File.WriteAllBytes(fileName, attach.Data);

                                // When we find an in-line attachment we have to replace the CID tag inside the HTML body
                                // with the name of the in-line attachment. But before we do this we check if the CID exists.
                                // When the CID does not exists we treat the in-line attachment as a normal attachment
                                if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
                                {
                                    body = body.Replace("cid:" + attach.ContentId, GetRelativePathFromAbsolutePath(fileName));
                                    continue;
                                }

                                result.Add(fileName);

                            }
                            else if (attachment.GetType() == typeof(Storage.Message))
                            {
                                var msg = (Storage.Message)attachment;
                                fileName =
                                    FileManager.FileExistsMakeNew(outputFolder +
                                                                  FileManager.RemoveInvalidFileNameChars(msg.Subject) +
                                                                  ".msg").Replace("&", "[and]");
                                result.Add(fileName);
                                msg.Save(fileName);
                            }

                            //var ext = Path.GetExtension(fileName);

                            if (attachments == string.Empty)
                                //attachments = BuildAnchor(fileName, !string.IsNullOrEmpty(ext) && ext.ToLower() == ".msg" ? MsgHandlerUrl : string.Empty); // Path.GetFileName(fileName);
                                attachments = BuildAnchor(fileName, MsgHandlerUrl);
                            else
                                //attachments += ", " + BuildAnchor(fileName, !string.IsNullOrEmpty(ext) && ext.ToLower() == ".msg" ? MsgHandlerUrl : string.Empty); // Path.GetFileName(fileName);
                                attachments += ", " + BuildAnchor(fileName, MsgHandlerUrl);
                        }

                        string outlookHeader;

                        if (htmlBody)
                        {
                            // Add an outlook style header into the HTML body.
                            // Change this code to the language you need. 
                            // Currently it is written in ENGLISH
                            outlookHeader =
                                "<TABLE cellSpacing=0 cellPadding=0 width=\"100%\" border=0 style=\"font-family: 'Times New Roman'; font-size: 12pt;\"\\>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>From:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + GetEmailSender(message) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>To:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + GetEmailRecipients(message, Storage.RecipientType.To) + "</TD></TR>" + Environment.NewLine +
                                "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Sent on:</STRONG></TD><TD valign=\"top\" style=\"height: 18px\">" + (message.SentOn != null ? ((DateTime)message.SentOn).ToString("dd-MM-yyyy HH:mm:ss") : string.Empty) + "</TD></TR>";

                            // CC
                            var cc = GetEmailRecipients(message, Storage.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>CC:</STRONG></TD><TD style=\"height: 18px\">" + cc + "</TD></TR>" + Environment.NewLine;

                            // Subject
                            outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Subject:</STRONG></TD><TD style=\"height: 18px\">" + message.Subject + "</TD></TR>" + Environment.NewLine;

                            // Empty line
                            outlookHeader += "<TR><TD colspan=\"2\" style=\"height: 18px\">&nbsp</TD></TR>" + Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader += "<TR><TD valign=\"top\" style=\"height: 18px; width: 100px \"><STRONG>Attachments:</STRONG></TD><TD style=\"height: 18px\">" + attachments + "</TD></TR>" + Environment.NewLine;

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
                                "Sent on:\t" + (message.SentOn != null ? ((DateTime)message.SentOn).ToString("dd-MM-yyyy HH:mm:ss") : string.Empty) + Environment.NewLine;

                            // CC
                            var cc = GetEmailRecipients(message, Storage.RecipientType.Cc);
                            if (cc != string.Empty)
                                outlookHeader += "CC:\t\t" + cc + Environment.NewLine;

                            outlookHeader += "Subject:\t" + message.Subject + Environment.NewLine + Environment.NewLine;

                            // Attachments
                            if (attachments != string.Empty)
                                outlookHeader += "Attachments:\t" + attachments + Environment.NewLine + Environment.NewLine;

                            body = outlookHeader + body;
                        }

                        // Write the body to a file
                        File.WriteAllText(eMailFileName, body);
                        return result.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                //if (message != null)
                //    message.Dispose();
                _errorMessage = GetInnerException(e);
                return new string[0];
            }
        }

        private string BuildAnchor(string fileFullName, string customHandler = "")
        {
            var fileName = Path.GetFileName(fileFullName);
            if (string.IsNullOrEmpty(customHandler))
            {
                return string.Format(@"<a href=""{0}"">{1}</a> ", GetRelativePathFromAbsolutePath(fileFullName, true), fileName);
            }
            return string.Format(@"<a href=""{0}"">{1}</a> ", customHandler + "&attachment=" + fileName, fileName);
        }

        private string GetRelativePathFromAbsolutePath(string path, bool encode = false)
        {
            var virtualDir = VirtualPath;
            virtualDir = virtualDir == "/" ? virtualDir : (virtualDir + "/");
            path = path.Replace(ApplicationPath, virtualDir).Replace(@"\", "/");
            return !encode ? path : HttpUtility.UrlEncode(path, System.Text.Encoding.UTF8);
        }
    }

    public static class MessageReaderFactory
    {
        public static IReader CreateMessageReader(HttpContext context)
        {
            return new WebReader(context.Request.RawUrl,
                                 context.Request.PhysicalApplicationPath,
                                 context.Request.ApplicationPath);
        }
    }
}

