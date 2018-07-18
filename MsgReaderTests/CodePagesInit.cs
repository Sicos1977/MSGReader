using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MsgReaderTests
{
    [TestClass]
    public class CodePagesInit
    {
        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
    }
}