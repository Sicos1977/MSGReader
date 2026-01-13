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
        public void MessageSizeProperty_UpdatedWhenExists()
        {
            // This test verifies that IF the PR_MESSAGE_SIZE property exists in the original file,
            // it gets updated after attachments are deleted and the file is saved.
            // Note: Not all MSG files have this property, so we test the update logic conditionally.
            
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
            
            // If the original file had PR_MESSAGE_SIZE property, verify it was updated
            if (originalSize.HasValue)
            {
                Assert.IsNotNull(newSize, "PR_MESSAGE_SIZE should still exist after saving");
                Assert.AreNotEqual(originalSize, newSize,
                    "PR_MESSAGE_SIZE should be updated (changed) after deleting attachments");
                    
                // The updated size should be reasonably close to the actual file size
                var percentageDifference = (double)System.Math.Abs((long)newSize - actualFileSize) / actualFileSize * 100;
                Assert.IsTrue(percentageDifference < 100, 
                    $"Updated PR_MESSAGE_SIZE ({newSize}) should be reasonably related to actual file size ({actualFileSize}). Difference: {percentageDifference:F2}%");
            }
            else
            {
                // If PR_MESSAGE_SIZE didn't exist originally, it's ok if it still doesn't exist
                // (Creating new properties is not currently supported)
                Assert.Inconclusive("PR_MESSAGE_SIZE property does not exist in the original file, cannot test update functionality");
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
