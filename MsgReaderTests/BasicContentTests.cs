using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;

namespace MsgReaderTests
{
    /*
         * Contains some basic content tests to make sure that the msg files are
         * being parsed correctly.  Sample emails contain chapter 1 from book 1
         * of War and Peace taken from Project Gutenberg.
         * 
         * Sample files are set to copy always content so will be accessible to
         * this project build location as relative path of SampleFiles/<file>.msg
         * 
         * MsgReader always returns HTML from .ExtractMsgEmailBody irrespective of
         * starting format, so all tests perform a simple tag cleanup regex.  This
         * will not work for all variations of HTML/XHTML, but will be fine for the simple
         * examples in this test.
         */

    [TestClass]
    public class BasicContentTests
    {
        private static readonly Regex _htmlSimpleCleanup = new Regex(@"<[^>]*>", RegexOptions.Compiled);
        private static readonly string _sampleText = "Heavens! what a virulent attack!";

        [TestMethod]
        public void Html_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\HtmlSampleEmail.msg"))
            {
                var msgReader = new Reader();
                var content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }
        }

        [TestMethod]
        public void Rtf_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\RtfSampleEmail.msg"))
            {
                var msgReader = new Reader();
                var content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }
        }

        [TestMethod]
        public void PlainText_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\TxtSampleEmail.msg"))
            {
                var msgReader = new Reader();
                var content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }
        }
    }
}