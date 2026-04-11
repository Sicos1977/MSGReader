using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System.IO;

namespace MsgReaderTests
{
    /// <summary>
    ///     Tests that verify sender email address extraction, including the fallback
    ///     to PR_SENDER_SMTP_ADDRESS_ALTERNATE (0x5D0A) for messages sent via Exchange.
    /// </summary>
    [TestClass]
    public class SenderTests
    {
        [TestMethod]
        public void Exchange_Sender_Email_NotEmpty()
        {
            // Arrange
            using var stream = File.OpenRead(Path.Combine("SampleFiles", "sender_not_found_from_exchange.msg"));

            // Act
            using var message = new Storage.Message(stream);

            // Assert – the sender e-mail must be populated for Exchange-originated messages
            Assert.IsNotNull(message.Sender, "Sender should not be null");
            Assert.IsFalse(string.IsNullOrEmpty(message.Sender.Email),
                "Sender.Email should not be empty for messages sent from Exchange server");
            Assert.IsTrue(message.Sender.Email.Contains("@"),
                $"Sender.Email '{message.Sender.Email}' should be a valid e-mail address");
        }
    }
}
