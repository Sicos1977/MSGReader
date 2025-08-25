using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Text;
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

        private static void Deal(string filePath, byte[] rtf)
        {
            if (_generateTestData)
                File.WriteAllBytes(filePath, rtf);
            else
                Assert.AreEqual(expected: Encoding.ASCII.GetString(File.ReadAllBytes(filePath)), actual: Encoding.ASCII.GetString(rtf));
        }
    }
}
