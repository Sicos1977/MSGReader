using System;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;

namespace MsgReaderTests
{
    [TestClass]
    public class RtfDecompressorTests
    {
        // ReSharper disable once InconsistentNaming
        private const bool _generateTestData = false;

        [TestMethod]
        public void LzFu()
        {
            var rtf = RtfDecompressor.DecompressRtf(File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "LZFu.bin")));
            Deal(Path.Combine("SampleFiles", "rtf", "LZFu.rtf"), rtf);
        }

        [TestMethod]
        public void Mela()
        {
            var rtf = RtfDecompressor.DecompressRtf(File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "MELA.bin")));
            Deal(Path.Combine("SampleFiles", "rtf", "MELA.rtf"), rtf);
        }

        /// <summary>
        ///     Regression test for MELA-format RTF where uncompressedSize is set equal to
        ///     compressedSize instead of the actual data length. Some MSG writers don't account
        ///     for the 12-byte header overhead (RawSize + CompType + CRC) when setting
        ///     uncompressedSize, causing it to exceed the available data by 12 bytes.
        ///     Before the fix, DecompressRtf would silently return an all-zero byte array.
        /// </summary>
        [TestMethod]
        public void MelaSizeMismatch()
        {
            // Build a dummy RTF payload
            var rtfContent = Encoding.ASCII.GetBytes(
                @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Helvetica;}}" +
                @"\uc1\pard\plain\f0\fs20 Hello world.\par }");

            // Build a MELA blob where uncompressedSize == compressedSize (the bug).
            // Per MS-OXRTFCP, compressedSize includes the 12-byte overhead for the
            // RawSize, CompType, and CRC fields. So compressedSize = rtfLen + 12.
            // The bug: uncompressedSize is set to compressedSize (rtfLen + 12)
            // instead of the actual data length (rtfLen). This means the decompressor
            // tries to copy 12 more bytes than are available after the header.
            var compressedSize = rtfContent.Length + 12;
            var uncompressedSize = compressedSize; // BUG: should be rtfContent.Length

            var blob = new byte[4 + compressedSize]; // 4 (compressedSize field) + compressedSize
            BitConverter.GetBytes(compressedSize).CopyTo(blob, 0);
            BitConverter.GetBytes(uncompressedSize).CopyTo(blob, 4);
            BitConverter.GetBytes(0x414C454D).CopyTo(blob, 8); // MELA magic
            BitConverter.GetBytes(0).CopyTo(blob, 12); // CRC (unused for MELA)
            Array.Copy(rtfContent, 0, blob, 16, rtfContent.Length);

            var result = RtfDecompressor.DecompressRtf(blob);

            var resultString = Encoding.ASCII.GetString(result);
            Assert.IsTrue(resultString.Contains(@"\rtf1"), "Decompressed output should contain valid RTF");
            Assert.IsTrue(resultString.Contains("Hello world."), "Decompressed output should contain the body text");
            Assert.AreEqual(rtfContent.Length, result.Length,
                "Result length should match actual data, not the inflated uncompressedSize");
        }

        private static void Deal(string filePath, byte[] rtf)
        {
            if (_generateTestData)
                File.WriteAllBytes(filePath, rtf);
            else
                Assert.AreEqual(expected: Encoding.ASCII.GetString(File.ReadAllBytes(filePath)), actual: Encoding.ASCII.GetString(rtf));
        }
    }
}
