
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
        public class RIGHT : TestBase
    {
            /// <summary>
            /// Passing in VBScript Empty as the string will return in a blank string being returned (so long as the length argument can be interpreted as a non-negative number)
            /// </summary>
            [TestMethod, MyFact]
            public void EmptyLengthOneReturnsBlankString()
            {
                myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT(null, 1));
            }

            /// <summary>
            /// Passing in VBScript Null as the string will return in VBScript Null being returned (so long as the length argument can be interpreted as a non-negative number)
            /// </summary>
            [TestMethod, MyFact]
            public void NullLengthOneReturnsNull()
            {
                myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT(DBNull.Value, 1));
            }

            [TestMethod, MyFact]
            public void ZeroLengthIsAcceptable()
            {
                myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("", 0));
            }

            [TestMethod, MyFact]
            public void NegativeLengthIsNotAcceptable()
            {
                myAssert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("", -1);
                });
            }

            [TestMethod, MyFact]
            public void EmptyLengthIsTreatedAsZeroLength()
            {
                myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abc", null));
            }

            [TestMethod, MyFact]
            public void MaxLengthLongerThanInputStringLengthIsTreatedAsEqualingInputStringLength()
            {
                myAssert.AreEqual("abc", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abc", 10));
            }

            [TestMethod, MyFact]
            public void NullLengthIsNotAcceptable()
            {
            myAssert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("", DBNull.Value);
                });
            }

            [TestMethod, MyFact]
            public void EnormousLengthResultsInOverflow()
            {
            myAssert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("", 1000000000000000);
                });
            }

            // These tests all illustrate that VBScript's standard "banker's rounding" is applied to fractional lengths
            [TestMethod, MyFact]
            public void LengthZeroPointFiveTreatedAsLengthZero()
            {
                myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 0.5));
            }
            [TestMethod, MyFact]
            public void LengthZeroPointNineTreatedAsLengthOne()
            {
                myAssert.AreEqual("d", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 0.9));
            }
            [TestMethod, MyFact]
            public void LengthOnePointFiveTreatedAsLengthTwo()
            {
                myAssert.AreEqual("cd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 1.5));
            }
            [TestMethod, MyFact]
            public void LengthOnePointNineTreatedAsLengthTwo()
            {
                myAssert.AreEqual("cd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 1.9));
            }
            [TestMethod, MyFact]
            public void LengthTwoPointFiveTreatedAsLengthTwo()
            {
                myAssert.AreEqual("cd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 2.5));
            }
            [TestMethod, MyFact]
            public void LengthTwoPointNineTreatedAsLengthThree()
            {
                myAssert.AreEqual("bcd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 2.9));
            }
            [TestMethod, MyFact]
            public void LengthThreePointFiveTreatedAsLengthFour()
            {
                myAssert.AreEqual("abcd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 3.5));
            }
            [TestMethod, MyFact]
            public void LengthThreePointNineTreatedAsLengthFour()
            {
                myAssert.AreEqual("abcd", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RIGHT("abcd", 3.9));
            }
        }
    //}
}
