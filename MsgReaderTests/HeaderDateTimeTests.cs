using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;

namespace MsgReaderTests
{
    [TestClass]
    public class HeaderDateTimeTests
    {
        private const string UtcFormat = "yyyy-MM-dd HH:mm:ss 'UTC'";
        private static readonly string MsgFile = Path.Combine("SampleFiles", "HtmlSampleEmail.msg");
        private static readonly string EmlFile = Path.Combine("SampleFiles", "TestWithAttachments.eml");

        [TestMethod]
        public void Msg_Default_UsesLocalTimeAndDefaultFormat()
        {
            using var message = new MsgReader.Outlook.Storage.Message(MsgFile);
            var expected = message.SentOn!.Value.ToString("F");

            var header = new Reader().ExtractMsgEmailHeader(message, ReaderHyperLinks.None);

            StringAssert.Contains(header, expected);
        }

        [TestMethod]
        public void Msg_HeaderTimeZoneAndFormat_AreUsedForSentOn()
        {
            using var message = new MsgReader.Outlook.Storage.Message(MsgFile);
            var expected = message.SentOn!.Value.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC";

            var reader = new Reader { HeaderTimeZone = TimeZoneInfo.Utc, HeaderDateTimeFormat = UtcFormat };
            var header = reader.ExtractMsgEmailHeader(message, ReaderHyperLinks.None);

            StringAssert.Contains(header, expected);
        }

        [TestMethod]
        public void Msg_HeaderTimeZone_ConvertsToTheGivenTimeZone()
        {
            var timeZone = TimeZoneInfo.CreateCustomTimeZone("MsgReaderTest+05:30", new TimeSpan(5, 30, 0), "Test", "Test");
            using var message = new MsgReader.Outlook.Storage.Message(MsgFile);
            var expected = message.SentOn!.Value.UtcDateTime.AddHours(5.5).ToString("yyyy-MM-dd HH:mm:ss") + " +05:30";

            var reader = new Reader { HeaderTimeZone = timeZone, HeaderDateTimeFormat = "yyyy-MM-dd HH:mm:ss zzz" };
            var header = reader.ExtractMsgEmailHeader(message, ReaderHyperLinks.None);

            StringAssert.Contains(header, expected);
        }

        [TestMethod]
        public void Msg_ExtractToFolder_UsesHeaderTimeZoneAndFormat()
        {
            DateTimeOffset sentOn;
            using (var message = new MsgReader.Outlook.Storage.Message(MsgFile))
                sentOn = message.SentOn!.Value;

            var outputFolder = CreateOutputFolder();
            try
            {
                var reader = new Reader { HeaderTimeZone = TimeZoneInfo.Utc, HeaderDateTimeFormat = UtcFormat };
                var files = reader.ExtractToFolder(MsgFile, outputFolder);
                var content = File.ReadAllText(files.First());

                StringAssert.Contains(content, sentOn.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC");
            }
            finally
            {
                Directory.Delete(outputFolder, true);
            }
        }

        [TestMethod]
        public void Eml_ExtractToFolder_UsesHeaderTimeZoneAndFormat()
        {
            var eml = MsgReader.Mime.Message.Load(new FileInfo(EmlFile));
            var expected = eml.Headers.DateSent.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC";

            var outputFolder = CreateOutputFolder();
            try
            {
                var reader = new Reader { HeaderTimeZone = TimeZoneInfo.Utc, HeaderDateTimeFormat = UtcFormat };
                var files = reader.ExtractToFolder(EmlFile, outputFolder);
                var content = File.ReadAllText(files.First());

                StringAssert.Contains(content, expected);
            }
            finally
            {
                Directory.Delete(outputFolder, true);
            }
        }

        private static string CreateOutputFolder()
        {
            var folder = Path.Combine(Path.GetTempPath(), "MsgReaderTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            return folder;
        }
    }
}
