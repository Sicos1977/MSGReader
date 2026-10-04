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

        /// <summary>
        ///     A double byte document code page (\ansicpg949, Korean) combined with a font table that mixes
        ///     \fcharset129 (Hangul) and \fcharset0 (ANSI). The Korean text is written under the \fcharset0 font,
        ///     which is what Outlook does in practice.
        /// </summary>
        /// <remarks>
        ///     The mixed font table plus a high byte under a single byte font encoding routes these bytes through
        ///     TryDecode. Charset detection cannot identify two bytes, so the fallback encoding decides the result.
        ///     Falling back to the single byte font encoding (cp1252) silently turns each double byte Korean
        ///     character into two Latin characters, so the document code page has to win here.
        /// </remarks>
        [TestMethod]
        public void MixedCharsetFontsWithDoubleByteDocumentCodePage()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg949\\fromhtml1" +
                "{\\fonttbl{\\f0\\fnil\\fcharset129 Malgun Gothic;}{\\f1\\fnil\\fcharset0 Calibri;}}" +
                "\\htmlrtf{\\htmlrtf0\\f1 \\'be\\'c8\\'b3\\'e7\\'c7\\'cf\\'bc\\'bc\\'bf\\'e4}}");
            Assert.AreEqual(expected: "안녕하세요", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        /// <summary>
        ///     A single byte document code page must not override the font encoding, so that a font declaring a
        ///     different single byte charset than the document keeps winning.
        /// </summary>
        [TestMethod]
        public void MixedCharsetFontsWithSingleByteDocumentCodePageKeepsFontEncoding()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg1252\\fromhtml1" +
                "{\\fonttbl{\\f0\\fnil\\fcharset204 Arial;}{\\f1\\fnil\\fcharset129 Malgun Gothic;}}" +
                "\\htmlrtf{\\htmlrtf0\\f0 \\'cf\\'f0\\'e8\\'e2\\'e5\\'f2}}");
            Assert.AreEqual(expected: "Привет", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        /// <summary>
        ///     Issue 529. A Japanese (\ansicpg932) mail where a &amp;nbsp; emits a \fcharset0 font inside its own
        ///     group, and the Japanese text on the following line inherits that single byte encoding.
        /// </summary>
        /// <remarks>
        ///     This is the structure Word produces for a NO-BREAK SPACE: {\f1\'a0} switches to a single byte ANSI
        ///     font, and the double byte Shift-JIS text that follows carries no font of its own. Decoding it with
        ///     cp1252 turns each character into two Latin characters, so \'82\'a0 becomes U+201A U+00A0 instead of
        ///     U+3042.
        /// </remarks>
        [TestMethod]
        public void MixedCharsetFontsWithNoBreakSpaceBeforeDoubleByteText()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg932\\fromhtml1" +
                "{\\fonttbl{\\f0\\fswiss\\fcharset128 MS PGothic;}{\\f1\\fmodern\\fcharset0 Courier New;}}" +
                "\\f0\\htmlrtf {\\f1\\'a0}\\htmlrtf0 \\htmlrtf {\\htmlrtf0 \\'82\\'a0\\'82\\'a0\\'82\\'a0}}");
            Assert.AreEqual(expected: "あああ", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        /// <summary>
        ///     A font that is selected inside a group must not stay active after the group ends. Outlook writes a
        ///     &amp;nbsp; as {\f1\'a0}, where \f1 is a single byte \fcharset0 font, in the middle of Simplified Chinese
        ///     (\fcharset134) text.
        /// </summary>
        /// <remarks>
        ///     The document declares the single byte code page \ansicpg1252, which is what Outlook writes when the
        ///     sender uses an English locale, so the document code page cannot be used to recover the text. When the
        ///     \f1 encoding leaks past the closing brace, the GBK bytes after the &amp;nbsp; are decoded as cp1252 and
        ///     中文 (\'d6\'d0\'ce\'c4) silently becomes ÖÐÎÄ.
        /// </remarks>
        [TestMethod]
        public void FontSelectedInsideGroupDoesNotLeakPastGroupEnd()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\deff0" +
                "{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\fmodern\\fcharset0 Courier New;}{\\f2\\fnil\\fcharset134 SimSun;}}" +
                "\\htmlrtf {\\f2 \\htmlrtf0 \\'d6\\'d0\\'ce\\'c4" +
                "{\\*\\htmltag84 &nbsp;}\\htmlrtf {\\f1\\'a0}\\htmlrtf0 \\'d6\\'d0\\'ce\\'c4" +
                "\\htmlrtf }\\htmlrtf0 }");
            Assert.AreEqual(expected: "中文&nbsp;中文", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        /// <summary>
        ///     The same as <see cref="FontSelectedInsideGroupDoesNotLeakPastGroupEnd" /> but with nested groups, the
        ///     encoding of each enclosing group has to be restored in turn.
        /// </summary>
        [TestMethod]
        public void FontSelectedInsideNestedGroupsIsRestoredPerGroup()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\deff0" +
                "{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\fnil\\fcharset204 Arial Cyr;}{\\f2\\fnil\\fcharset134 SimSun;}}" +
                "\\htmlrtf {\\f2 \\htmlrtf0 \\'d6\\'d0\\'ce\\'c4 " +
                "\\htmlrtf {\\f1 \\htmlrtf0 \\'cf\\'f0\\'e8\\'e2\\'e5\\'f2 \\htmlrtf {\\f0 \\htmlrtf0 caf\\'e9\\htmlrtf }\\htmlrtf0  \\'cf\\'f0\\'e8\\'e2\\'e5\\'f2\\htmlrtf }\\htmlrtf0 " +
                " \\'d6\\'d0\\'ce\\'c4\\htmlrtf }\\htmlrtf0 }");
            Assert.AreEqual(expected: "中文 Привет café Привет 中文", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        /// <summary>
        ///     High bytes that are read before any font has been selected must be decoded with the document code page
        ///     instead of throwing.
        /// </summary>
        [TestMethod]
        public void HighBytesBeforeAnyFontIsSelectedUseDocumentCodePage()
        {
            var rtfDomDocument = new Document();
            rtfDomDocument.DeEncapsulateHtmlFromRtf(
                "{\\rtf1\\ansi\\ansicpg1252\\fromhtml1" +
                "{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\fnil\\fcharset134 SimSun;}}" +
                "\\htmlrtf{\\htmlrtf0 caf\\'e9 na\\'efve}}");
            Assert.AreEqual(expected: "café naïve", actual: rtfDomDocument.HtmlContent, ignoreCase: false);
        }

        private static void Deal(string filePath, string rtf)
        {
            Assert.AreEqual(expected: File.ReadAllText(filePath), actual: rtf);
        }
    }
}
