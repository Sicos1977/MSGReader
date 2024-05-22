using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

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
    }
}