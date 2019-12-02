
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;

namespace MsgReaderTests
{
    [TestClass]
    public class RemoveAttachmentTests
    {
        [TestMethod]
        public void RemoveAttachments()
        {
            using (var inputStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg")))
            using (var inputMessage = new Storage.Message(inputStream, FileAccess.ReadWrite))
            {
                var attachments = inputMessage.Attachments.ToList();

                foreach (var attachment in attachments)
                    inputMessage.DeleteAttachment(attachment);

                using (var outputStream = new MemoryStream())
                {
                    inputMessage.Save(outputStream);
                    using (var outputMessage = new Storage.Message(outputStream))
                    {
                        var count = outputMessage.Attachments.Count;
                        Assert.IsTrue(count == 0);
                    }
                }
            }
        }
    }
}
