using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text;

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