using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Lib.Strings;

namespace Tests.Lib.Strings
{
    /// <summary>
    ///     Test class for function "StringToUtf8Bytes"
    /// </summary>
    [TestClass]
    public class StringToUtf8Bytes
    {
        [TestMethod]
        public void TestWithRegularString1()
        {
            byte[] result = StringsFunctions.StringToUtf8Bytes("Test1");

            CollectionAssert.AreEqual(
                result, 
                new byte[] { 84, 101, 115, 116, 49 }
            );
        }

        [TestMethod]
        public void TestWithRegularString2()
        {
            byte[] result = StringsFunctions.StringToUtf8Bytes("Hack The Planet!!!");

            CollectionAssert.AreEqual(
                result,
                new byte[] { 72, 97, 99, 107, 32, 84, 104, 101, 32, 80, 108, 97, 110, 101, 116, 33, 33, 33 }
            );
        }

        [TestMethod]
        public void TestWithEmptyString()
        {
            byte[] result = StringsFunctions.StringToUtf8Bytes(string.Empty);

            CollectionAssert.AreEqual(
                result,
                null
            );
        }

        [TestMethod]
        public void TestWithNullString()
        {
            byte[] result = StringsFunctions.StringToUtf8Bytes(null);

            CollectionAssert.AreEqual(
                result,
                null
            );
        }

        [TestMethod]
        public void TestWithOneCharString()
        {
            byte[] result = StringsFunctions.StringToUtf8Bytes("0");

            CollectionAssert.AreEqual(
                result,
                new byte[] { 48 }
            );
        }
    }
}
