using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System.IO;

namespace MsgReaderTests
{
    /*
     * Contains some basic tests to make sure that reactions are exporting
     * correctly.
     * 
     * Sample files are set to copy always content so will be accessible to
     * this project build location as relative path of SampleFiles/<file>.msg
     */

    [TestClass]
    public class ReactionTests
    {
        private const string ReactionSummaryRowText = "test2, test2@readreceipts.onmicrosoft.com, heart, 10/19/2023 22:36:00";
        private const string ReactionSummaryLabelText = "Reactions summary";
        private const string OwnerReactionHistoryRowText = "test2, test2@readreceipts.onmicrosoft.com, [removed], 10/19/2023 22:35:25";
        private const string OwnerReactionHistoryLabelText = "Owner reaction history";

        [TestMethod]
        public void Reactions_ExtractMsgEmailBody()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "EmailWithReactions.msg"));

            var reader = new Reader();
            var content = reader.ExtractMsgEmailBody(fileStream, ReaderHyperLinks.Both, null, includeReactionsInfo: true);

            Assert.IsNotNull(content);
            Assert.IsTrue(content.Contains(ReactionSummaryRowText));
            Assert.IsTrue(content.Contains(ReactionSummaryLabelText));
            Assert.IsTrue(content.Contains(OwnerReactionHistoryRowText));
            Assert.IsTrue(content.Contains(OwnerReactionHistoryLabelText));
        }

        [TestMethod]
        public void Reactions_ExtractToFolder()
        {
            var inputFile = Path.Combine("SampleFiles", "EmailWithReactions.msg");

            var reader = new Reader();
            var tempDir = GetTempDir();
            var files = reader.ExtractToFolder(inputFile, tempDir, includeReactionsInfo: true);

            Directory.Delete(tempDir, true);
        }

        [TestMethod]
        public void Reactions_ExtractToFolder_NoReactions()
        {
            var inputFile = Path.Combine("SampleFiles", "HtmlSampleEmailWithAttachment.msg");

            var reader = new Reader();
            var tempDir = GetTempDir();
            var files = reader.ExtractToFolder(inputFile, tempDir, includeReactionsInfo: true);

            Directory.Delete(tempDir, true);
        }

        [TestMethod]
        public void Reactions_ExtractMsgEmailBody_NoReactions()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "HtmlSampleEmailWithAttachment.msg"));

            var reader = new Reader();
            var content = reader.ExtractMsgEmailBody(fileStream, ReaderHyperLinks.Both, null, includeReactionsInfo: true);

            Assert.IsNotNull(content);
            Assert.IsFalse(content.Contains(ReactionSummaryRowText));
        }

        private static string GetTempDir()
        {
            var temp = Path.GetTempFileName();
            File.Delete(temp);
            Directory.CreateDirectory(temp);
            return temp;
        }
    }
}