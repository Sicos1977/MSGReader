using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace MsgReaderTests
{

    /*
     * Contains some basic tests to make sure that attachments are exporting
     * correctly.
     * 
     * Sample files are set to copy always content so will be accessible to
     * this project build location as relative path of SampleFiles/<file>.msg
     * 
     * The attachment for sample files is known and the SHA1 is hardcoded in
     * this test class.  A valid extraction should yield an exact copy of
     * the attached file.
     * 
     * Tests in this class write to local temp files for extraction, using
     * a random file name (which is guaranteed unique) and uses that as a
     * directory name by deleting the new file and creating a dir in its place
     */

    [TestClass]
    public class AttachmentTests
    {
        private static string _knownSHA1 = "600F7BED593588956BFC3CFEEFEC12506120B653";

        [TestMethod]
        public void Html_Single_Attachment_Test()
        {
            Reader reader = new Reader();
            string tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder("SampleFiles\\HtmlSampleEmailWithAttachment.msg", tempDirPath);

            List<string> sha1s = new List<string>();

            foreach (string filePath in outputFiles)
            {
                sha1s.Add(GetSha1(filePath));
            }

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1s.Contains(_knownSHA1));
        }

        [TestMethod]
        public void Rtf_Single_Attachment_Test()
        {
            Reader reader = new Reader();
            string tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder("SampleFiles\\RtfSampleEmailWithAttachment.msg", tempDirPath);

            List<string> sha1s = new List<string>();

            foreach (string filePath in outputFiles)
            {
                sha1s.Add(GetSha1(filePath));
            }

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1s.Contains(_knownSHA1));
        }

        [TestMethod]
        public void Txt_Single_Attachment_Test()
        {
            Reader reader = new Reader();
            string tempDirPath = GetTempDir();
            IEnumerable<string> outputFiles = reader.ExtractToFolder("SampleFiles\\TxtSampleEmailWithAttachment.msg", tempDirPath);

            List<string> sha1s = new List<string>();

            foreach (string filePath in outputFiles)
            {
                sha1s.Add(GetSha1(filePath));
            }

            Directory.Delete(tempDirPath, true);

            Assert.IsTrue(sha1s.Contains(_knownSHA1));
        }

        private static string GetTempDir()
        {
            string temp = Path.GetTempFileName();
            File.Delete(temp);
            Directory.CreateDirectory(temp);
            return temp;
        }

        private static string GetSha1(string filePath)
        {
            using (SHA1 sha1 = SHA1.Create())
            using (Stream stream = File.OpenRead(filePath))
            {
                byte[] hashBytes = sha1.ComputeHash(stream);

                StringBuilder sb = new StringBuilder();

                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }

                return sb.ToString();
            }
        }
    }
}
