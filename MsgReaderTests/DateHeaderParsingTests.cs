using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Mime.Header;
using System.Diagnostics;

namespace MsgReaderTests
{
    [TestClass]
    public class DateHeaderParsingTests
    {
        [TestMethod]
        public void DateHeader_WithTimeZone_IsParsedCorrectly()
        {
            // Arrange: sample EML content with a Date header
            string eml = @"Date: Fri, 2 May 2025 12:27:25 +0000
MIME-Version: 1.0
To: some@example.com
From: someone@example.com
Subject: Test

Body text.";

            // Act: extract headers and parse
            var headers = HeaderExtractor.GetHeaders(eml);

            // Output the parsed result for verification
            Debug.WriteLine("Parsed DateSent: " + headers.DateSent.ToString("o"));

            // Assert: the DateSent should be in UTC (+0000)
            Assert.AreEqual(2025, headers.DateSent.Year);
            Assert.AreEqual(5, headers.DateSent.Month);
            Assert.AreEqual(2, headers.DateSent.Day);
            Assert.AreEqual(12, headers.DateSent.Hour);
            Assert.AreEqual(0, headers.DateSent.Offset.Hours); // +0000
        }

        [TestMethod]
        public void DateHeader_WithDifferentTimeZone_IsParsedCorrectly()
        {
            // Arrange: sample EML content with a Date header in +0200
            string eml = @"Date: Fri, 2 May 2025 14:27:25 +0200
MIME-Version: 1.0
To: some@example.com
From: someone@example.com
Subject: Test

Body text.";

            // Act: extract headers and parse
            var headers = HeaderExtractor.GetHeaders(eml);

            // Output the parsed result for verification
            System.Diagnostics.Debug.WriteLine("Parsed DateSent: " + headers.DateSent.ToString("o"));

            // Assert: the DateSent should be in +0200
            Assert.AreEqual(2025, headers.DateSent.Year);
            Assert.AreEqual(5, headers.DateSent.Month);
            Assert.AreEqual(2, headers.DateSent.Day);
            Assert.AreEqual(14, headers.DateSent.Hour);
            Assert.AreEqual(2, headers.DateSent.Offset.Hours); // +0200
        }

        [TestMethod]
        public void DateHeader_WithNegativeTimeZone_IsParsedCorrectly()
        {
            // Arrange: sample EML content with a Date header in -0500
            string eml = @"Date: Fri, 2 May 2025 07:27:25 -0500
MIME-Version: 1.0
To: some@example.com
From: someone@example.com
Subject: Test

Body text.";

            // Act: extract headers and parse
            var headers = HeaderExtractor.GetHeaders(eml);

            // Output the parsed result for verification
            Debug.WriteLine("Parsed DateSent: " + headers.DateSent.ToString("o"));

            // Assert: the DateSent should be in -0500
            Assert.AreEqual(2025, headers.DateSent.Year);
            Assert.AreEqual(5, headers.DateSent.Month);
            Assert.AreEqual(2, headers.DateSent.Day);
            Assert.AreEqual(7, headers.DateSent.Hour);
            Assert.AreEqual(-5, headers.DateSent.Offset.Hours); // -0500
        }
    }

}