using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Lib.Strings;

namespace Tests.Lib.Strings
{
    /// <summary>
    ///     Test function BytesAsHexString
    /// </summary>
    [TestClass]
    public class BytesAsHexString
    {
        [TestMethod]
        public void BytesAsHexStringWithEmptyString()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { }),
                ""
            );
        }

        [TestMethod]
        public void BytesAsHexStringWithZero()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 0 }),
                "00"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith1Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 66 }),
                "42"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith3Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 1, 13, 255 }),
                "010DFF"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith7Numbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 0, 4, 1, 13, 255, 254, 88 }),
                "0004010DFFFE58"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith8Zeros()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0 }),
                "0000000000000000"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith9LargeNumbers()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 255, 255, 255, 255, 255, 255, 255, 255, 255 }),
                "FFFFFFFFFFFFFFFFFF"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith1NumberAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 66 }, ";"),
                "42"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith3NumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 1, 13, 255 }, ","),
                "01,0D,FF"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith7NumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 0, 4, 1, 13, 255, 254, 88 }, "-"),
                "00-04-01-0D-FF-FE-58"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith8ZerosAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0 }, " $"),
                "00 $00 $00 $00 $00 $00 $00 $00"
            );
        }

        [TestMethod]
        public void BytesAsHexStringWith9LargeNumbersAndSplitter()
        {
            Assert.AreEqual(
                StringsFunctions.BytesAsHexString(new byte[] { 255, 255, 255, 255, 255, 255, 255, 255, 255 }, "&"),
                "FF&FF&FF&FF&FF&FF&FF&FF&FF"
            );
        }
    }
}
