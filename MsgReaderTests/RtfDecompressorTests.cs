using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System.IO;
using System.Text;

namespace MsgReaderTests
{
    [TestClass]
    public class RtfDecompressorTests
    {
        // ReSharper disable once InconsistentNaming
        private static readonly bool _generateTestData = false;

        [TestMethod]
        public void LzFu()
        {
            var rtf = RtfDecompressor.DecompressRtf(
                File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "LZFu.bin"))
            );
            Deal(Path.Combine("SampleFiles", "rtf", "LZFu.rtf"), rtf);
        }

        [TestMethod]
        public void Mela()
        {
            var rtf = RtfDecompressor.DecompressRtf(
                File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "Mela.bin"))
            );
            Deal(Path.Combine("SampleFiles", "rtf", "Mela.rtf"), rtf);
        }

        private void Deal(string filePath, byte[] rtf)
        {
            if (_generateTestData)
            {
                File.WriteAllBytes(
                    filePath,
                    rtf
                );
            }
            else
            {
                Assert.AreEqual(
                    expected: Encoding.ASCII.GetString(File.ReadAllBytes(filePath)),
                    actual: Encoding.ASCII.GetString(rtf)
                );
            }
        }
    }
}
