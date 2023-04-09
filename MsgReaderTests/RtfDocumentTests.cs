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
        [TestMethod]
        public void ParseTextF1()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.ParseRtfText("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\htmlrtf{\\lang1030 \\htmlrtf0  f\\'f8}}");
            Assert.AreEqual(expected: " fø", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        [TestMethod]
        public void ParseTextF2()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.ParseRtfText("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\htmlrtf{\\lang1030 \\htmlrtf0 f\\'f8}}");
            Assert.AreEqual(expected: "fø", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }
    }
}
