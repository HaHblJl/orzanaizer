using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace трпо
{
    [TestClass]
    public class трпоTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            decimal x = 1000;
            decimal y = 2000;
            decimal expected = 1500;
            трпо.Sred sr = new трпо.Sred();
            decimal actual = sr.Srd(x,y);
            Assert.AreEqual(expected, actual);
        }
    }
}
