using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace BlueColor.TestLog.Tests
{
    [TestClass]
    public class HelloWorldTest
    {
        [TestMethod]
        public void SayHelloTest()
        {
            string reuslt = HelloWorld.SayHello("BlueLover");

            Assert.AreEqual(reuslt, "Hello,BlueLover!");
        }
    }
}