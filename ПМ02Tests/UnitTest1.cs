using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ПМ_02;

namespace ПМ02Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            double money = 100;
            double count = 0;
            double expend = 0;

            Form1 form1 = new Form1();
            double actual = 0;

            Assert.AreEqual(expend, actual);
        }
        [TestMethod]
        public void TestMethod2()
        {

            double expend = 0;

            Form1 form1 = new Form1();
            double actual = 0;

            Assert.AreEqual(expend, actual);
        }
        [TestMethod]
        public void TestMethod3()
        {

            double expend = 0;

            Form1 form1 = new Form1();
            double actual = 0;

            Assert.AreEqual(expend, actual);
        }
    }
}
