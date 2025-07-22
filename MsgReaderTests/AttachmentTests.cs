using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using MsgReader.Mime;

namespace MsgReaderTests
{
    /*
     * Contains some basic tests to make sure that attachments are exporting
     * correctly.
     * 
     * Sample files are set to copy always content so will be accessible to
     * this project build location as relative path of SampleFiles/<file>.msg
     * 
     * The attachment for sample files is known and the SHA1 is hardcoded in
     * this test class.  A valid extraction should yield an exact copy of
     * the attached file.
     * 
     * Tests in this class write to local temp files for extraction, using
     * a random file name (which is guaranteed unique) and uses that as a
     * directory name by deleting the new file and creating a dir in its place
     */

    [TestClass]
    public class AttachmentTests
    {
        private const string KnownSha1 = "600F7BED593588956BFC3CFEEFEC12506120B653";

        [TestMethod]
        public void Html_Single_Attachment_Test()
        {
            var reader = new Reader();
            var tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder(Path.Combine("SampleFiles", "HtmlSampleEmailWithAttachment.msg"), tempDirPath);

            var sha1S = outputFiles.Select(GetSha1).ToList();

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1S.Contains(KnownSha1));
        }

        [TestMethod]
        public void Rtf_Single_Attachment_Test()
        {
            var reader = new Reader();
            var tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder(Path.Combine("SampleFiles", "RtfSampleEmailWithAttachment.msg"), tempDirPath);

            var sha1S = outputFiles.Select(GetSha1).ToList();

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1S.Contains(KnownSha1));
        }

        [TestMethod]
        public void Txt_Single_Attachment_Test()
        {
            var reader = new Reader();
            var tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder(Path.Combine("SampleFiles", "TxtSampleEmailWithAttachment.msg"), tempDirPath);

            var sha1S = outputFiles.Select(GetSha1).ToList();

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1S.Contains(KnownSha1));
        }

        [TestMethod]
        public void Load_From_Message_Constructor_Test()
        {
            var message = new MsgReader.Outlook.Storage.Message(new FileInfo(Path.Combine("SampleFiles", "TxtSampleEmailWithAttachment.msg")).OpenRead());
            Assert.IsTrue(message.Attachments.Any());
        }

        [TestMethod]
        public void Remove_Attachments_From_EML_Test()
        {
            const string fileName = "TestWithAttachmentsAndInlines";

            var eml = LoadEmlWithAttachments(fileName);
            foreach (var attachment in eml.Attachments.Where(a => a.IsAttachment && !a.IsInline).ToArray())
            {
                eml.Attachments.Remove(attachment);
            }

            var outputFileName = $"{fileName}_out";
            eml.Save(BuildFileInfo(outputFileName));

            eml = LoadEmlWithAttachments(outputFileName);

            Assert.AreEqual(0, eml.Attachments.Count);
        }

        [TestMethod]
        public void Mail_with_inline_with_ContentDisposition_Test()
        {
            const string fileName = "TestWithInlineHavingContentDisposition";
            var outputFileName = $"{fileName}_out";
            
            var eml = LoadEmlWithAttachments(fileName);
            eml.Attachments.RemoveAt(0);
            eml.Save(BuildFileInfo(outputFileName));

            eml = LoadEmlWithAttachments(outputFileName);

            Assert.AreEqual(0, eml.Attachments.Count);
        }

        [TestMethod]
        public void Mail_with_inline_without_ContentDisposition_Test()
        {
            const string fileName = "TestWithInlineHavingNoContentDisposition";
            var outputFileName = $"{fileName}_out";

            var eml = LoadEmlWithAttachments(fileName);
            eml.Attachments.RemoveAt(0);
            eml.Save(BuildFileInfo(outputFileName));

            eml = LoadEmlWithAttachments(outputFileName);
            
            Assert.AreEqual(0, eml.Attachments.Count);
        }

        [TestMethod]
        public void Mail_not_loading_attachments_Test()
        {
            const string fileName = "TestWithAttachmentsAndInlines";
            var outputFileName = $"{fileName}_out";

            var eml = LoadEmlWithoutAttachments(fileName);
            eml.Save(BuildFileInfo(outputFileName));

            eml = LoadEmlWithoutAttachments(outputFileName);
            
            Assert.AreEqual(4, eml.Attachments.Count);

            var emlAttachment = eml.Attachments[0];
            Assert.IsTrue(emlAttachment.IsAttachment);
            Assert.IsFalse(emlAttachment.IsInline);
            Assert.IsFalse(emlAttachment.IsMultiPart);
            Assert.AreEqual(0, emlAttachment.Body.Length);
            Assert.IsFalse(emlAttachment.ContentDisposition.Inline);
            Assert.AreEqual("attachment", emlAttachment.ContentDisposition.DispositionType);
            Assert.AreEqual("attachment1.txt", emlAttachment.ContentDisposition.FileName);

            emlAttachment = eml.Attachments[1];
            Assert.IsTrue(emlAttachment.IsAttachment);
            Assert.IsFalse(emlAttachment.IsInline);
            Assert.IsFalse(emlAttachment.IsMultiPart);
            Assert.AreEqual(0, emlAttachment.Body.Length);
            Assert.IsFalse(emlAttachment.ContentDisposition.Inline);
            Assert.AreEqual("attachment", emlAttachment.ContentDisposition.DispositionType);
            Assert.AreEqual("attachment2.txt", emlAttachment.ContentDisposition.FileName);

            emlAttachment = eml.Attachments[2];
            Assert.IsTrue(emlAttachment.IsAttachment);
            Assert.IsFalse(emlAttachment.IsInline);
            Assert.IsFalse(emlAttachment.IsMultiPart);
            Assert.AreEqual(0, emlAttachment.Body.Length);
            Assert.IsFalse(emlAttachment.ContentDisposition.Inline);
            Assert.AreEqual("attachment", emlAttachment.ContentDisposition.DispositionType);
            Assert.AreEqual("attachment3.txt", emlAttachment.ContentDisposition.FileName);

            emlAttachment = eml.Attachments[3];
            Assert.IsTrue(emlAttachment.IsAttachment);
            Assert.IsFalse(emlAttachment.IsInline);
            Assert.IsFalse(emlAttachment.IsMultiPart);
            Assert.AreEqual(0, emlAttachment.Body.Length);
            Assert.IsFalse(emlAttachment.ContentDisposition.Inline);
            Assert.AreEqual("attachment", emlAttachment.ContentDisposition.DispositionType);
            Assert.AreEqual("attachment4.txt", emlAttachment.ContentDisposition.FileName);
        }

        [TestMethod]
        public void Msg_with_msg_attachment_Test()
        {
            var message = new MsgReader.Outlook.Storage.Message(new FileInfo(Path.Combine("SampleFiles", "EmailWithMsgAttachment.msg")).OpenRead());
            Assert.HasCount(1, message.Attachments);
            Assert.IsInstanceOfType<MsgReader.Outlook.Storage.Message>(message.Attachments[0]);
            var innerMessage = (MsgReader.Outlook.Storage.Message)message.Attachments[0];
            using var ms = new MemoryStream();
            innerMessage.Save(ms);
            Assert.AreEqual(58880, ms.Length);
        }

        private static FileInfo BuildFileInfo(string fileName)
        {
            return new FileInfo(Path.Combine("SampleFiles", $"{fileName}.eml"));
        }

        private static Message LoadEmlWithAttachments(string fileName)
        {
            Message eml;
            using (var fileStream = BuildFileInfo(fileName).OpenRead())
            {
                eml = Message.Load(fileStream, true, true);
            }

            return eml;
        }

        private static Message LoadEmlWithoutAttachments(string fileName)
        {
            Message eml;
            using (var fileStream = BuildFileInfo(fileName).OpenRead())
            {
                eml = Message.Load(fileStream, true, false);
            }

            return eml;
        }

        private static string GetTempDir()
        {
            var temp = Path.GetTempFileName();
            File.Delete(temp);
            Directory.CreateDirectory(temp);
            return temp;
        }

        private static string GetSha1(string filePath)
        {
            using var sha1 = SHA1.Create();
            using Stream stream = File.OpenRead(filePath);
            var hashBytes = sha1.ComputeHash(stream);

            var sb = new StringBuilder();

            foreach (var t in hashBytes)
                sb.Append(t.ToString("X2"));

            return sb.ToString();
        }
    }
}