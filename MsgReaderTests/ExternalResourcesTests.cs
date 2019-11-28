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
            int? SampleOperationWithDotNetsStreamReadersFinalizer(System.IO.Stream stream)
            {
                //We will not call dispose here so the finalizer should not kill the extern stream
                var reference = new System.IO.StreamReader(stream);
                return reference.Read();
            }

            int? OperationWithMsgReaderFinalizer(System.IO.Stream stream)
            {
                //We will not call dispose here so the finalizer should not kill the extern stream
                var reference = new Storage.Message(stream);
                return reference.Size;
            }

            using (var inputStream = System.IO.File.Open(System.IO.Path.Combine("SampleFiles", "EmailWithAttachments.msg"), System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                SampleOperationWithDotNetsStreamReadersFinalizer(inputStream);
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

                OperationWithMsgReaderFinalizer(inputStream);
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
