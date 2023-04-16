using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MsgReaderTests
{
    [TestClass]
    public class RtfDecompressorTests
    {
        private static bool _generateTestData = false;

        [TestMethod]
        public void LZFu()
        {
            var rtf = RtfDecompressor.DecompressRtf(
                File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "LZFu.bin"))
            );
            Deal(Path.Combine("SampleFiles", "rtf", "LZFu.rtf"), rtf);
        }

        [TestMethod]
        public void MELA()
        {
            var rtf = RtfDecompressor.DecompressRtf(
                File.ReadAllBytes(Path.Combine("SampleFiles", "rtf", "MELA.bin"))
            );
            Deal(Path.Combine("SampleFiles", "rtf", "MELA.rtf"), rtf);
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
