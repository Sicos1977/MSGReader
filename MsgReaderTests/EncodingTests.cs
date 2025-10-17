using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using MsgReader.Mime.Decode;

namespace MsgReaderTests
{
    [TestClass]
    public class EncodingTests
    {
        [TestMethod]
        public void Subject_Special_Chars()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithSpecialCharsInSubject.msg"));
            {
                using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
                Assert.IsTrue(msg.Subject == "Un sujet bien défini Àéroport mañana être ouïe électricité così próxima à Ô");
            }
        }

        [TestMethod]
        public void Subject_Special_Chars_2()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithSpecialCharsInSubject_2.msg"));
            {
                using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
                Assert.IsTrue(msg.Subject == "Un sujet très bien défini");
            }
        }

        [TestMethod]
        public void From_ISO_8859_1()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithISO_8859_1_From.eml"));
            var msg = new MsgReader.Mime.Message(fileStream);
            Assert.IsTrue(msg.Headers.From.Address == "some@example.com");
            Assert.IsTrue(msg.Headers.From.DisplayName == "Some ßtring");
        }

        [TestMethod]
        public void Encoded_And_Unencoded()
        {
            var decoded = EncodedWord.Decode("=?iso-8859-1?Q?Some_=DFtring?= <some@example.com>");
            Assert.IsTrue(decoded == "Some ßtring <some@example.com>");
        }

        [TestMethod]
        public void Encoded_And_Unencoded_2()
        {
            var decoded = EncodedWord.Decode("=?iso-8859-1?Q?Some_=DFtring?= <some@example.com> =?iso-8859-1?Q?Some_=DFtring?=");
            Assert.IsTrue(decoded == "Some ßtring <some@example.com> Some ßtring");
        }

        [TestMethod]
        public void Multiline_Subject_With_Leading_Space()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithMultilineSubject.eml"));
            var msg = new MsgReader.Mime.Message(fileStream);
            // RFC 2822 section 2.2.3: When unfolding headers, leading whitespace should be preserved as a single space
            Assert.AreEqual("One Two Three Four Five Six Seven Eight Nine Ten Eleven Twelve Thirteen Fourteen Fifteen", msg.Headers.Subject);
            // Verify the space between "Twelve" and "Thirteen" is present
            Assert.IsTrue(msg.Headers.Subject.Contains("Twelve Thirteen"));
        }

        [TestMethod]
        public void Attachment_Inherits_Parent_Codepage()
        {
            // Test that attachments inherit the parent message's codepage for proper decoding
            // of PT_STRING8 properties like attachment filenames with special characters
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg"));
            {
                using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
                
                // Verify that the message has attachments
                Assert.IsTrue(msg.Attachments.Count > 0, "Message should have attachments");
                
                // Verify that attachments can access their filenames without encoding errors
                foreach (var attachment in msg.Attachments)
                {
                    if (attachment is MsgReader.Outlook.Storage.Attachment att)
                    {
                        // Filename should not contain replacement characters (�)
                        Assert.IsFalse(att.FileName.Contains('\ufffd'), 
                            $"Attachment filename '{att.FileName}' should not contain Unicode replacement character");
                        Assert.IsFalse(string.IsNullOrEmpty(att.FileName), 
                            "Attachment filename should not be empty");
                    }
                }
            }
        }
    }
}