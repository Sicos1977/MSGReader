
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System.IO;
using System.Linq;

namespace MsgReaderTests
{
    [TestClass]
    public class RemoveAttachmentTests
    {
        [TestMethod]
        public void RemoveAttachments()
        {
            using var inputStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg"));
            using var inputMessage = new Storage.Message(inputStream, FileAccess.ReadWrite);
            var attachments = inputMessage.Attachments.ToList();

            foreach (var attachment in attachments)
                inputMessage.DeleteAttachment(attachment);

            using var outputStream = new MemoryStream();
            inputMessage.Save(outputStream);
            using var outputMessage = new Storage.Message(outputStream);
            var count = outputMessage.Attachments.Count;
            Assert.IsTrue(count == 0);
        }

        [TestMethod]
        public void RemoveAttachmentsFromInner()
        {
            using var memoryStream = new MemoryStream();
            using var inputStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithInnerMailAndAttachments.msg"));
            inputStream.CopyTo(memoryStream);
            memoryStream.Position = 0;
            using var inputMessage = new Storage.Message(memoryStream, FileAccess.ReadWrite);
            Assert.AreEqual(3, inputMessage.Attachments.Count);
            Assert.AreEqual("OUTER 1.pdf", ((Storage.Attachment)inputMessage.Attachments[0]).FileName);
            Assert.AreEqual("OUTER 2.pdf", ((Storage.Attachment)inputMessage.Attachments[1]).FileName);
            Assert.AreEqual("Inner mail.msg", ((Storage.Message)inputMessage.Attachments[2]).FileName);

            var innerMessage1 = (Storage.Message)inputMessage.Attachments[2];

            {
                var innerAttachments1 = innerMessage1.Attachments.ToList();

                Assert.AreEqual(2, innerAttachments1.Count);
                Assert.AreEqual("INNER 1.pdf", ((Storage.Attachment)innerAttachments1[0]).FileName);
                Assert.AreEqual("INNER 2.pdf", ((Storage.Attachment)innerAttachments1[1]).FileName);

                // Delete `INNER 2.pdf` from inner mail.

                innerMessage1.DeleteAttachment(innerAttachments1[1]);
            }

            var innerAttachments2 = innerMessage1.Attachments.ToList();

            // Ensure `INNER 2.pdf` has been gone.

            Assert.AreEqual(1, innerAttachments2.Count);
            Assert.AreEqual("INNER 1.pdf", ((Storage.Attachment)innerAttachments2[0]).FileName);

            using var outputStream = new MemoryStream();
            inputMessage.Save(outputStream);
            using var outputMessage = new Storage.Message(outputStream);
            Assert.AreEqual(3, outputMessage.Attachments.Count); // <-- Assert.AreEqual failed. Expected:<3>. Actual:<2>. 
            Assert.AreEqual("OUTER 1.pdf", ((Storage.Attachment)outputMessage.Attachments[0]).FileName);
            Assert.AreEqual("OUTER 2.pdf", ((Storage.Attachment)outputMessage.Attachments[1]).FileName);
            Assert.AreEqual("Inner mail.msg", ((Storage.Message)outputMessage.Attachments[2]).FileName);

            var innerMessage2 = (Storage.Message)outputMessage.Attachments[2];

            var innerAttachments = innerMessage2.Attachments.ToList();

            Assert.AreEqual(1, innerAttachments2.Count);
            Assert.AreEqual("INNER 1.pdf", ((Storage.Attachment)innerAttachments[0]).FileName);
        }
    }
}
