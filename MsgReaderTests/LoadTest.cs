using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System.IO;
using MsgReader.Mime.Decode;

namespace MsgReaderTests
{
    [TestClass]
    public class LoadTest
    {
        private DirectoryInfo _tempDirectory;

        [TestInitialize]
        public void Initialize()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            _tempDirectory = Directory.CreateDirectory(tempDirectory);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _tempDirectory.Delete(true);
        }

        [TestMethod]
        [Ignore("Long running stress test - takes 55+ seconds. Should be run separately for performance testing.")]
        public void Extract_10000_Times()
        {
            for (var i = 0; i < 10000; i++)
            {
                var msgReader = new Reader();
                var tempDirectory = Directory.CreateDirectory(Path.Combine(_tempDirectory.FullName, Path.GetRandomFileName()));
                msgReader.ExtractToFolder(Path.Combine("SampleFiles", "HtmlSampleEmail.msg"), tempDirectory.FullName);
            }
        }
    }
}