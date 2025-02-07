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
    }
}