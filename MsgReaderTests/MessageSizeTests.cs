using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System.IO;
using System.Linq;

namespace MsgReaderTests
{
    [TestClass]
    public class MessageSizeTests
    {
        [TestMethod]
        public void ActualFileSize_DecreasesAfterAttachmentDeletion()
        {
            // Arrange
            using var inputStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg"));
            var originalFileSize = inputStream.Length;
            
            using var inputMessage = new Storage.Message(inputStream, FileAccess.ReadWrite);
            var originalSize = inputMessage.Size;
            var originalAttachmentCount = inputMessage.Attachments.Count;
            
            Assert.AreEqual(2, originalAttachmentCount, "Original message should have 2 attachments");
            
            // Act - Delete all attachments
            var attachments = inputMessage.Attachments.ToList();
            foreach (var attachment in attachments)
                inputMessage.DeleteAttachment(attachment);
            
            // Save to new stream
            using var outputStream = new MemoryStream();
            inputMessage.Save(outputStream);
            
            var actualFileSize = outputStream.Length;
            
            // Reload the saved message
            outputStream.Position = 0;
            using var savedMessage = new Storage.Message(outputStream);
            
            var newSize = savedMessage.Size;
            var newAttachmentCount = savedMessage.Attachments.Count;
            
            // Assert
            Assert.AreEqual(0, newAttachmentCount, "Saved message should have 0 attachments");
            
            // The file should be significantly smaller after removing attachments (at least 20% smaller)
            Assert.IsTrue(actualFileSize < originalFileSize * 0.8, 
                $"Actual file size ({actualFileSize}) should be at least 20% less than original file size ({originalFileSize})");
            
            // If PR_MESSAGE_SIZE existed originally and exists after saving, it should be updated
            if (originalSize.HasValue && newSize.HasValue)
            {
                Assert.IsTrue(newSize < originalSize, 
                    $"New PR_MESSAGE_SIZE ({newSize}) should be less than original ({originalSize}) when property exists");
                    
                // The PR_MESSAGE_SIZE should also be reasonably close to the actual file size
                var percentageDifference = (double)System.Math.Abs((long)newSize - actualFileSize) / actualFileSize * 100;
                Assert.IsTrue(percentageDifference < 50, 
                    $"PR_MESSAGE_SIZE ({newSize}) should be within 50% of actual file size ({actualFileSize}). Difference: {percentageDifference:F2}%");
            }
        }

        [TestMethod]
        public void AttachmentsRemovedSuccessfully()
        {
            // Arrange
            using var inputStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWith2Attachments.msg"));
            using var inputMessage = new Storage.Message(inputStream, FileAccess.ReadWrite);
            
            // Act - Delete all attachments
            var attachments = inputMessage.Attachments.ToList();
            foreach (var attachment in attachments)
                inputMessage.DeleteAttachment(attachment);
            
            // Save to new stream
            using var outputStream = new MemoryStream();
            inputMessage.Save(outputStream);
            
            // Reload the saved message
            outputStream.Position = 0;
            using var savedMessage = new Storage.Message(outputStream);
            
            // Assert - Attachments should be gone
            Assert.AreEqual(0, savedMessage.Attachments.Count, 
                "Saved message should have 0 attachments after deletion");
        }
    }
}
