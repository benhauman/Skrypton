

using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class LEFT : TestBase
    {
        /// <summary>
        /// Passing in VBScript Empty as the string will return in a blank string being returned (so long as the length argument can be interpreted as a non-negative number)
        /// </summary>
        [TestMethod]
        public void EmptyLengthOneReturnsBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT(null, 1));
        }

        /// <summary>
        /// Passing in VBScript Null as the string will return in VBScript Null being returned (so long as the length argument can be interpreted as a non-negative number)
        /// </summary>
        [TestMethod]
        public void NullLengthOneReturnsNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().LEFT(DBNull.Value, 1));
        }

        [TestMethod]
        public void ZeroLengthIsAcceptable()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("", 0));
        }

        [TestMethod]
        public void NegativeLengthIsNotAcceptable()
        {
            myAssert.Throws<InvalidProcedureCallOrArgumentException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("", -1);
            });
        }

        [TestMethod]
        public void EmptyLengthIsTreatedAsZeroLength()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abc", null));
        }

        [TestMethod]
        public void NullLengthIsNotAcceptable()
        {
            myAssert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("", DBNull.Value);
                });
        }

        [TestMethod]
        public void MaxLengthLongerThanInputStringLengthIsTreatedAsEqualingInputStringLength()
        {
            myAssert.AreEqual("abc", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abc", 10));
        }

        [TestMethod]
        public void EnormousLengthResultsInOverflow()
        {
            myAssert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("", 1000000000000000);
                });
        }

        // These tests all illustrate that VBScript's standard "banker's rounding" is applied to fractional lengths
        [TestMethod]
        public void LengthZeroPointFiveTreatedAsLengthZero()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 0.5));
        }
        [TestMethod]
        public void LengthZeroPointNineTreatedAsLengthOne()
        {
            myAssert.AreEqual("a", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 0.9));
        }
        [TestMethod]
        public void LengthOnePointFiveTreatedAsLengthTwo()
        {
            myAssert.AreEqual("ab", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 1.5));
        }
        [TestMethod]
        public void LengthOnePointNineTreatedAsLengthTwo()
        {
            myAssert.AreEqual("ab", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 1.9));
        }
        [TestMethod]
        public void LengthTwoPointFiveTreatedAsLengthTwo()
        {
            myAssert.AreEqual("ab", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 2.5));
        }
        [TestMethod]
        public void LengthTwoPointNineTreatedAsLengthThree()
        {
            myAssert.AreEqual("abc", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 2.9));
        }
        [TestMethod]
        public void LengthThreePointFiveTreatedAsLengthFour()
        {
            myAssert.AreEqual("abcd", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 3.5));
        }
        [TestMethod]
        public void LengthThreePointNineTreatedAsLengthFour()
        {
            myAssert.AreEqual("abcd", DefaultRuntimeSupportClassFactoryInstance.Get().LEFT("abcd", 3.9));
        }
    }
    //}
}
