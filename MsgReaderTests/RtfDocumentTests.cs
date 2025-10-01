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
        public void ParseLineBreak()
        {
            // Test that \line is converted to newline
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\htmlrtf{\\htmlrtf0 line1\\line line2\\line line3}}");
            Assert.AreEqual(expected: "line1\r\nline2\r\nline3", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        [TestMethod]
        public void ParseLineBreakInPreTag()
        {
            // Test that \line inside <pre> tags preserves the newlines
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1 {\\*\\htmltag128 <pre>}\\htmlrtf{\\htmlrtf0  com.sun.mail.smtp.SMTPSendFailedException: 451\\line \\line \\tab  at com.sun.mail.smtp.SMTPTransport.sendMessage\\line }}");
            Assert.IsTrue(rtfDomDocument.HtmlContent.Contains("\r\n"), "HTML content should contain newlines from \\line");
        }

        [TestMethod]
        public void Issue332()
        {
            var rtfDomDocument = new Document();
            // RTF files should be read as ASCII/Latin1 since RTF control words are ASCII
            var rtfContent = File.ReadAllText("SampleFiles/rtf/Issue332.rtf", System.Text.Encoding.GetEncoding("ISO-8859-1"));
            rtfDomDocument.DeEncapsulateHtmlFromRtf(rtfContent);
            Deal("SampleFiles/rtf/Issue332.html", rtfDomDocument.HtmlContent);
        }

        /// <summary>
        /// Using CP932 (Shift_JIS) in rtf
        /// </summary>
        [TestMethod]
        public void Issue347()
        {
            var rtfDomDocument = new Document();
            // RTF files should be read as ASCII/Latin1 since RTF control words are ASCII
            var rtfContent = File.ReadAllText("SampleFiles/rtf/Issue347.rtf", System.Text.Encoding.GetEncoding("ISO-8859-1"));
            rtfDomDocument.DeEncapsulateHtmlFromRtf(rtfContent);
            Deal("SampleFiles/rtf/Issue347.html", rtfDomDocument.HtmlContent);
        }

        private static void Deal(string filePath, string rtf)
        {
            Assert.AreEqual(expected: File.ReadAllText(filePath), actual: rtf);
        }
    }
}
