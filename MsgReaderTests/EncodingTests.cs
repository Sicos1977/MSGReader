using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using System.Text;
using MsgReader.Mime.Decode;
#pragma warning disable MSTEST0037

namespace MsgReaderTests
{
    [TestClass]
    public class EncodingTests
    {
        [TestMethod]
        public void Subject_Special_Chars()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithSpecialCharsInSubject.msg"));
            using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
            Assert.IsTrue(msg.Subject == "Un sujet bien défini Àéroport mañana être ouïe électricité così próxima à Ô");
        }

        [TestMethod]
        public void Subject_Special_Chars_2()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithSpecialCharsInSubject_2.msg"));
            using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
            Assert.IsTrue(msg.Subject == "Un sujet très bien défini");
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
        public void Subject_GB2312_MultiBlock_Base64()
        {
            // Verifies that multi-block Base64-encoded subjects (e.g. GB2312) are decoded correctly.
            // Each block must be Base64-decoded independently before the results are concatenated.
            var encoded =
                "=?GB2312?B?SVMgUEFUUklDSUEgKERXSCkgqaZSZXBvcnR5IGZlYnJ1qKJyIDIwMjYsIA==?=" +
                "=?GB2312?B?T2Jkb2JpZTogMDEuMDEuMjAyNi0yOC4wMi4yMDI2LCBIaXN0qK5yaWEgaw==?=" +
                "=?GB2312?B?IDAzLjAzLjIwMjY=?=";
            var decoded = EncodedWord.Decode(encoded);
            Assert.AreEqual("IS PATRICIA (DWH) \u2502Reporty febru\u00e1r 2026, Obdobie: 01.01.2026-28.02.2026, Hist\u00f3ria k 03.03.2026", decoded);
        }

        [TestMethod]
        public void Attachment_Inherits_Parent_Codepage()
        {
            // Test that attachments inherit the parent message's codepage for proper decoding
            // of PT_STRING8 properties like attachment filenames with special characters
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg"));
            using var msg = new MsgReader.Outlook.Storage.Message(fileStream);
            
            // Verify that the message has attachments
            Assert.IsTrue(msg.Attachments.Count > 0, "Message should have attachments");
            
            // Verify that attachments can access their filenames without encoding errors
            foreach (var attachment in msg.Attachments)
            {
                if (attachment is not MsgReader.Outlook.Storage.Attachment att) continue;
                // Filename should not contain replacement characters (�)
                Assert.IsFalse(att.FileName.Contains('\ufffd'), 
                    $"Attachment filename '{att.FileName}' should not contain Unicode replacement character");
                Assert.IsFalse(string.IsNullOrEmpty(att.FileName), 
                    "Attachment filename should not be empty");
            }
        }

        [TestMethod]
        public void DecodeString8_Falls_Back_To_Current_Culture_Ansi_Codepage_When_Utf8_Is_Garbled()
        {
            var expected = "测试样本中文附件名称占位用例数据内容补齐字段.docx";
            var bytes = Encoding.GetEncoding(936).GetBytes(expected);

            var decoded = MsgReader.Outlook.Storage.DecodeString8(bytes, Encoding.UTF8, Encoding.GetEncoding(936));

            Assert.AreEqual(expected, decoded);
        }

        [TestMethod]
        public void DecodeString8_Keeps_Declared_Codepage_When_Text_Is_Already_Valid()
        {
            var expected = "测试样本中文附件名称占位用例数据内容补齐字段.docx";
            var bytes = Encoding.UTF8.GetBytes(expected);

            var decoded = MsgReader.Outlook.Storage.DecodeString8(bytes, Encoding.UTF8, Encoding.GetEncoding(936));

            Assert.AreEqual(expected, decoded);
        }

        [TestMethod]
        public void Contains_Ansi_Body_Stream_Returns_True_When_Present()
        {
            var entries = new[]
            {
                "__substg1.0_001A001F",
                "__substg1.0_1000001E"
            };

            Assert.IsTrue(MsgReader.Outlook.Storage.ContainsAnsiBodyStream(entries));
        }

        [TestMethod]
        public void Contains_Ansi_Body_Stream_Is_Case_Insensitive()
        {
            var entries = new[]
            {
                "__substg1.0_001A001F",
                "__SUBSTG1.0_1000001e"
            };

            Assert.IsTrue(MsgReader.Outlook.Storage.ContainsAnsiBodyStream(entries));
        }

        [TestMethod]
        public void Contains_Ansi_Body_Stream_Returns_False_When_Absent()
        {
            var entries = new[]
            {
                "__substg1.0_001A001F",
                "__substg1.0_1000001F"
            };

            Assert.IsFalse(MsgReader.Outlook.Storage.ContainsAnsiBodyStream(entries));
        }
    }
}