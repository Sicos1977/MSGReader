﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
        private static Regex _htmlSimpleCleanup = new Regex(@"<[^>]*>", RegexOptions.Compiled);
        private static string _sampleText = "Heavens! what a virulent attack!";

        [TestMethod]
        public void Html_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\HtmlSampleEmail.msg"))
            {
                Reader msgReader = new Reader();
                string content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }            
        }

        [TestMethod]
        public void Rtf_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\RtfSampleEmail.msg"))
            {
                Reader msgReader = new Reader();
                string content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }
        }

        [TestMethod]
        public void PlainText_Content_Test()
        {
            using (Stream fileStream = File.OpenRead("SampleFiles\\TxtSampleEmail.msg"))
            {
                Reader msgReader = new Reader();
                string content = msgReader.ExtractMsgEmailBody(fileStream, true, null);
                content = _htmlSimpleCleanup.Replace(content, string.Empty);
                Assert.IsTrue(content.Contains(_sampleText));
            }
        }
    }
}
