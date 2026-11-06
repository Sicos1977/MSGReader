using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Exceptions;
using MsgReader.Outlook;
using OpenMcdf;

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

        [TestMethod]
        public void InvalidHeadersAndSizesThrowRtfDecompressionException()
        {
            Assert.AreEqual("Invalid compressed-RTF header",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(null)).Message);
            Assert.AreEqual("Invalid compressed-RTF header",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(new byte[15])).Message);

            foreach (var size in new[] { -1, 100 * 1024 * 1024 + 1 })
            {
                var blob = CreateBlob(0x414c454d, size);
                var exception = Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(blob));
                StringAssert.Contains(exception.Message, "Invalid compressed-RTF uncompressed size");
            }

            var unknown = CreateBlob(0x69736e61, 0);
            Assert.AreEqual("Unknown compression type (magic number 1769172577)",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(unknown)).Message);
        }

        [TestMethod]
        public void InvalidPayloadsThrowRtfDecompressionException()
        {
            var mela = CreateBlob(0x414c454d, 1);
            Assert.AreEqual("uncompressed-RTF data size mismatch",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(mela)).Message);

            var lzfu = CreateBlob(0x75465a4c, 1);
            BitConverter.GetBytes(0).CopyTo(lzfu, 0);
            Assert.AreEqual("compressed-RTF data size mismatch",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(lzfu)).Message);

            lzfu = CreateBlob(0x75465a4c, 1);
            BitConverter.GetBytes(1).CopyTo(lzfu, 12);
            Assert.AreEqual("compressed-RTF CRC32 failed",
                Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(lzfu)).Message);

            foreach (var payload in new[] { Array.Empty<byte>(), new byte[] { 0 }, new byte[] { 1, 0 }, new byte[] { 1, 0xff, 0xf0 }, new byte[] { 1, 0, 0 } })
            {
                lzfu = CreateBlob(0x75465a4c, 1, payload);
                Assert.AreEqual("Invalid compressed-RTF data",
                    Assert.ThrowsExactly<RtfDecompressionException>(() => RtfDecompressor.DecompressRtf(lzfu)).Message);
            }
        }

        [TestMethod]
        public void BodyHtmlPropagatesRtfDecompressionException()
        {
            using var stream = new MemoryStream();
            using (var input = File.OpenRead(Path.Combine("SampleFiles", "RtfSampleEmail.msg")))
                input.CopyTo(stream);
            stream.Position = 0;
            using (var root = RootStorage.Open(stream, StorageModeFlags.LeaveOpen))
            {
                foreach (var entry in root.EnumerateEntries().ToList())
                    if (entry.Name.StartsWith("__substg1.0_1013", StringComparison.Ordinal) ||
                        entry.Name == "__substg1.0_10090102")
                        root.Delete(entry.Name);

                using (var rtf = root.CreateStream("__substg1.0_10090102"))
                    rtf.Write(CreateBlob(0x69736e61, 0));
                root.Flush();
            }

            stream.Position = 0;
            Assert.ThrowsExactly<RtfDecompressionException>(() =>
            {
                using var message = new MsgReader.Outlook.Storage.Message(stream);
                _ = message.BodyHtml;
            });
        }

        private static byte[] CreateBlob(int magic, int uncompressedSize, byte[] payload = null)
        {
            payload ??= Array.Empty<byte>();
            var blob = new byte[16 + payload.Length];
            BitConverter.GetBytes(blob.Length - 4).CopyTo(blob, 0);
            BitConverter.GetBytes(uncompressedSize).CopyTo(blob, 4);
            BitConverter.GetBytes(magic).CopyTo(blob, 8);
            uint crc = 0;
            foreach (var value in payload)
            {
                crc ^= value;
                for (var bit = 0; bit < 8; bit++)
                    crc = (crc >> 1) ^ ((crc & 1) != 0 ? 0xedb88320u : 0);
            }
            BitConverter.GetBytes(crc).CopyTo(blob, 12);
            payload.CopyTo(blob, 16);
            return blob;
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
