using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;

namespace MsgReaderTests
{
    [TestClass]
    public class ExternalResourcesTests
    {
        [TestMethod]
        [DeploymentItem("SampleFiles", "SampleFiles")]
        public void StorageFinaizerBehaviourTest()
        {
            int? SampleOperationWithDotNetsStreamReaders(System.IO.Stream stream)
            {
                //We will not call dispose here to keep the extern stream alive
                var reference = new System.IO.StreamReader(stream);
                return reference.Read();
            }

            int? OperationWithMsgReader(System.IO.Stream stream)
            {
                //We will not call dispose here to keep the extern stream alive
                var reference = new Storage.Message(stream);
                return reference.Size;
            }

            using (var inputStream = System.IO.File.Open(System.IO.Path.Combine("SampleFiles", "EmailWithAttachments.msg"), System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                SampleOperationWithDotNetsStreamReaders(inputStream);
                GC.Collect(0, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();
                try
                {
                    inputStream.Seek(0, System.IO.SeekOrigin.Begin);
                }
                catch (ObjectDisposedException)
                {
                    Assert.Fail("The stream should not be disposed now");
                }

                OperationWithMsgReader(inputStream);
                GC.Collect(0, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();
                try
                {
                    inputStream.Seek(0, System.IO.SeekOrigin.Begin);
                }
                catch (ObjectDisposedException)
                {
                    Assert.Fail("The stream should not be disposed now");
                }
            }
        }
    }
}
