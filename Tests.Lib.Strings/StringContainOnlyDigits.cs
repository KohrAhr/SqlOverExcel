using System;
using System.Text;
using Lib.Strings;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests.Lib.Strings
{
    [TestClass]
    public class StringContainOnlyDigits
    {
        [TestMethod]
        public void StringIsEmpty()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(""), false);
        }

        [TestMethod]
        public void StringIsNull()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(null), false);
        }

        [TestMethod]
        public void StringIsMaxValue()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(Char.MaxValue.ToString()), false);
        }

        [TestMethod]
        public void StringIsMinValue()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(Char.MinValue.ToString()), false);
        }

        [TestMethod]
        public void StringWithOnlyAlphabet()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("AbCdEf"), false);
        }

        [TestMethod]
        public void StringWithAlphabetAndNumbers()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("Ab34Cd54Ef"), false);
        }

        [TestMethod]
        public void StringWithAlphabetAndLeadingNumbers()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("0909Af"), false);
        }

        [TestMethod]
        public void StringWithAlphabetAndEndingNumbers()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("iwut0909"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndSpaces()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("12 345"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndEndingSpace()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("12345 "), false);
        }

        [TestMethod]
        public void StringWithNumbersAndLeadingSpace()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(" 3421"), false);
        }

        [TestMethod]
        public void StringWithNumbers()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("3421"), true);
        }

        [TestMethod]
        public void StringWithNumber()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("1"), true);
        }

        [TestMethod]
        public void StringWithNumbersAndSpecialChar()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("34!21"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndLeadingSpecialChar()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("?3421"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndEndingSpecialChar()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("3421%"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndComma()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("342,1"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndLeadingComma()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(",3421"), false);
        }


        [TestMethod]
        public void StringWithNumbersAndEndingComma()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("3421,"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndDot()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("34.21"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndLeadingDot()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits(".3421"), false);
        }

        [TestMethod]
        public void StringWithNumbersAndEndingDot()
        {
            Assert.AreEqual(StringsFunctions.StringContainOnlyDigits("3421."), false);
        }
    }
}
