using System;
using System.Text;
using Lib.Strings;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests.Lib.Strings
{
    [TestClass]
    public class BytesAsString
    {
        [TestMethod]
        public void BytesAsStringWithEmptyString()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { }), 
                ""
            );
        }

        [TestMethod]
        public void BytesAsStringWithZero()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 0 }),
                "0"
            );
        }

        [TestMethod]
        public void BytesAsStringWith1Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 66 }), 
                "66"
            );
        }

        [TestMethod]
        public void BytesAsStringWith3Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 1, 13, 255 }), 
                "113255"
            );
        }

        [TestMethod]
        public void BytesAsStringWith7Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 0, 4, 1, 13, 255, 254, 88 }),
                "0411325525488"
            );
        }

        [TestMethod]
        public void BytesAsStringWith8Zeros()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0 }),
                "00000000"
            );
        }

        [TestMethod]
        public void BytesAsStringWith9LargeNumbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 255, 255, 255, 255, 255, 255, 255, 255, 255 }),
                "255255255255255255255255255"
            );
        }

        [TestMethod]
        public void BytesAsStringWith1NumberAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 45 }, "@"),
                "45"
            );
        }

        [TestMethod]
        public void BytesAsStringWith3NumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 1, 13, 255 }, "+"),
                "1+13+255"
            );
        }

        [TestMethod]
        public void BytesAsStringWith7NumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 0, 4, 1, 13, 255, 254, 88 }, "~"),
                "0~4~1~13~255~254~88"
            );
        }

        [TestMethod]
        public void BytesAsStringWith8ZerosAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0 }, "&"),
                "0&0&0&0&0&0&0&0"
            );
        }

        [TestMethod]
        public void BytesAsStringWith9LargeNumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsString(new byte[] { 255, 255, 255, 255, 255, 255, 255, 255, 255 }, "^^^"),
                "255^^^255^^^255^^^255^^^255^^^255^^^255^^^255^^^255"
            );
        }
    }
}
