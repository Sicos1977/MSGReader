using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Rtf;
using System.IO;

namespace MsgReaderTests
{
    [TestClass]
    public class RtfDocumentTests
    {
        [TestMethod]
        public void ParseTextF1()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\htmlrtf{\\lang1030 \\htmlrtf0  f\\'f8}}");
            Assert.AreEqual(expected: " fø", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        [TestMethod]
        public void ParseTextF2()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\htmlrtf{\\lang1030 \\htmlrtf0 f\\'f8}}");
            Assert.AreEqual(expected: "fø", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        [TestMethod]
        public void Issue332()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(File.ReadAllText("SampleFiles/rtf/Issue332.rtf"));
            Deal("SampleFiles/rtf/Issue332.html", rtfDomDocument.HtmlContent);
        }

        /// <summary>
        /// Using CP932 (Shift_JIS) in rtf
        /// </summary>
        [TestMethod]
        public void Issue347()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(File.ReadAllText("SampleFiles/rtf/Issue347.rtf"));
            Deal("SampleFiles/rtf/Issue347.html", rtfDomDocument.HtmlContent);
        }

        private static void Deal(string filePath, string rtf)
        {
            Assert.AreEqual(expected: File.ReadAllText(filePath), actual: rtf);
        }
    }
}
