using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Rtf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MsgReaderTests
{
    [TestClass]
    public class RtfDocumentTests
    {
        private static readonly bool _generateTestData = false;

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

        private void Deal(string filePath, string rtf)
        {
            if (_generateTestData)
            {
                File.WriteAllText(
                    filePath,
                    rtf
                );
            }
            else
            {
                Assert.AreEqual(
                    expected: File.ReadAllText(filePath),
                    actual: rtf
                );
            }
        }
    }
}
