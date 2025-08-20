using Microsoft.VisualStudio.TestTools.UnitTesting;
using MsgReader.Outlook;
using System;
using System.Text;

namespace MsgReaderTests
{
    [TestClass]
    public class ReactionHelperTests
    {
        [TestMethod]
        public void GetReactionsFromOwnerReactionsHistory_Deserializes()
        {
            var json = "[{\"name\":\"A\",\"email\":\"a@example.com\",\"type\":\"like\",\"datetime\":\"2023-10-19T22:36:00+00:00\"}]";
            var bytes = Encoding.UTF8.GetBytes(json);
            var reactions = ReactionHelper.GetReactionsFromOwnerReactionsHistory(bytes);
            Assert.AreEqual(1, reactions.Count);
            Assert.AreEqual("A", reactions[0].Name);
            Assert.AreEqual("like", reactions[0].Type);
        }

        [TestMethod]
        public void GetReactionsFromOwnerReactionsHistory_InvalidJson_Throws()
        {
            var bytes = Encoding.UTF8.GetBytes("invalid");
            Assert.ThrowsException<ArgumentException>(() => ReactionHelper.GetReactionsFromOwnerReactionsHistory(bytes));
        }
    }
}
