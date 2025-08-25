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
        private static readonly Regex HtmlSimpleCleanup = new Regex(@"<[^>]*>", RegexOptions.Compiled);

        [TestMethod]
        public void Content_Test()
        {
            using Stream fileStream = File.OpenRead(Path.Combine("SampleFiles", "UTF-8_Test.eml"));
            var message = new MsgReader.Mime.Message(fileStream);
            string content = "";
            if (message.HtmlBody != null)
                content = System.Text.Encoding.UTF8.GetString(message.HtmlBody.Body);
            else if (message.TextBody != null)
                content = System.Text.Encoding.UTF8.GetString(message.TextBody.Body);
            content = HtmlSimpleCleanup.Replace(content, string.Empty);
            Assert.IsTrue(content.Contains(Body));
        }

        [TestMethod]
        public void FromName_Test()
        {
            var fileInfo = new FileInfo(Path.Combine("SampleFiles", "UTF-8_Test.eml"));
            var message = Message.Load(fileInfo);
            var from = message.Headers.From.DisplayName;
            // Debug: Check what we actually got
            System.Diagnostics.Debug.WriteLine($"Expected FromName: '{FromName}'");
            System.Diagnostics.Debug.WriteLine($"Actual DisplayName: '{from}'");
            System.Diagnostics.Debug.WriteLine($"From Address: '{message.Headers.From.Address}'");
            Assert.IsTrue(from != null && from.Contains(FromName), $"Expected '{FromName}' but got '{from}'");
        }

        [TestMethod]
        public void Subject_Test()
        {
            var fileInfo = new FileInfo(Path.Combine("SampleFiles", "UTF-8_Test.eml"));
            var message = Message.Load(fileInfo);
            var subject = message.Headers.Subject;
            // Debug: Check what we actually got
            System.Diagnostics.Debug.WriteLine($"Expected Subject: '{Subject}'");
            System.Diagnostics.Debug.WriteLine($"Actual Subject: '{subject}'");
            Assert.IsTrue(subject != null && subject.Contains(Subject), $"Expected '{Subject}' but got '{subject}'");
        }
    }
}