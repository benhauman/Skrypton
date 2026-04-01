
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
    //{
        public class CDBL : TestBase
    {
            [TestMethod]
            public void Empty()
            {
                myAssert.AreEqual(
                    0d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(null)
                );
            }

            [TestMethod]
            public void Null()
            {
                myAssert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(DBNull.Value);
                });
            }

            [TestMethod]
            public void BlankString()
            {
                myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL("");
                });
            }

            [TestMethod]
            public void NonNumericString()
            {
                myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL("a");
                });
            }

            [TestMethod]
            public void PositiveNumberAsString()
            {
                myAssert.AreEqual(
                    123.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL("123.4")
                );
            }

            [TestMethod]
            public void PositiveNumberAsStringWithLeadingAndTrailingWhitespace()
            {
                myAssert.AreEqual(
                    123.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(" 123.4 ")
                );
            }

            [TestMethod]
            public void PositiveNumberWithNoZeroBeforeDecimalPoint()
            {
                myAssert.AreEqual(
                    0.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(" .4 ")
                );
            }

            [TestMethod]
            public void NegativeNumberWithNoZeroBeforeDecimalPoint()
            {
                myAssert.AreEqual(
                    -0.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(" -.4 ")
                );
            }

            [TestMethod]
            public void NegativeNumberWithNoZeroBeforeDecimalPointAndSpaceBetweenSignAndPoint()
            {
                myAssert.AreEqual(
                    -0.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(" - .4 ")
                );
            }

            [TestMethod]
            public void NegativeNumberAsString()
            {
                myAssert.AreEqual(
                    -123.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL("-123.4")
                );
            }

            [TestMethod]
            public void Nothing()
            {
                var nothing = VBScriptConstants.Nothing;
                myAssert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(nothing);
                });
            }

            [TestMethod]
            public void ObjectWithoutDefaultProperty()
            {
                myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(new object());
                });
            }

            [TestMethod]
            public void ObjectWithDefaultProperty()
            {
                var target = new exampledefaultpropertytype { result = 123.4 };
                myAssert.AreEqual(
                    123.4,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(target)
                );
            }

            [TestMethod]
            public void Zero()
            {
                myAssert.AreEqual(
                    0d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(0)
                );
            }

            [TestMethod]
            public void PlusOne()
            {
                myAssert.AreEqual(
                    1d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(1)
                );
            }

            [TestMethod]
            public void MinusOne()
            {
                myAssert.AreEqual(
                    -1d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(-1)
                );
            }

            [TestMethod]
            public void OnePointOne()
            {
                myAssert.AreEqual(
                    1.1d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(1.1)
                );
            }

            [TestMethod]
            public void DateAndTime()
            {
                myAssert.AreEqualX(
                    42026.8410300926d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(new DateTime(2015, 1, 22, 20, 11, 5, 0)),
                    10 // This test fails without specifying precision
                );
            }

            [TestMethod]
            public void True()
            {
                myAssert.AreEqual(
                    -1d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(true)
                );
            }

            [TestMethod]
            public void False()
            {
                myAssert.AreEqual(
                    0d,
                    DefaultRuntimeSupportClassFactoryInstance.Get().CDBL(false)
                );
            }
        }
    //}
}
