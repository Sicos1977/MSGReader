using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System.IO;
using System.Text.RegularExpressions;
using MsgReader.Mime;

namespace MsgReaderTests
{
    [TestClass]
    public class UTF8Tests
    {
        private static readonly string FromName = "設備予約systemシステム";


        private static readonly string Subject = "【要確認】●普通自転車（ノーパンクタイヤ） 再予約通知 2023/11/29 12:00 - 13:00";


        private static readonly string Body = "メッセージです。通常通りパースされることを確認済みです。";

        [TestMethod]
        public void Content_Test()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "UTF-8_Test.eml"));
            var msgReader = new Reader();
            var content = msgReader.ExtractMsgEmailBody(fileStream, ReaderHyperLinks.Both, null);
            content = HtmlSimpleCleanup.Replace(content, string.Empty);
            Assert.IsTrue(content.Contains(Body));
        }

        [TestMethod]
        public void FromName_Test()
        {
            using var fileInfo = new DirectoryInfo(Path.Combine("SampleFiles", "UTF-8_Test.eml")).GetFiles()[0];
            var message = Message.Load(fileInfo);
            var from = message.Headers.From.DisplayName;
            Assert.IsTrue(from.Contains(FromName));
        }

        [TestMethod]
        public void Subject_Test()
        {
            using var fileInfo = new DirectoryInfo(Path.Combine("SampleFiles", "UTF-8_Test.eml")).GetFiles()[0];
            var message = Message.Load(fileInfo);
            var subject = message.Headers.Subject;
            Assert.IsTrue(subject.Contains(Subject));
        }
    }
}