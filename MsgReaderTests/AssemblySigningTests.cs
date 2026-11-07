using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;

namespace MsgReaderTests
{
    [TestClass]
    public class AssemblySigningTests
    {
        [TestMethod]
        public void MsgReaderHasExpectedStrongName()
        {
            var token = typeof(Storage).Assembly.GetName().GetPublicKeyToken();
            Assert.IsNotNull(token);
            Assert.AreEqual("3B-A0-1C-F3-43-49-59-FA", BitConverter.ToString(token));
        }
    }
}
