

using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Globalization;

//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
    //{
        public class NUM : TestBase
    {
            [TestMethod, MyFact]
            public void Empty()
            {
                myAssert.AreEqual(
                    (Int16)0,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(null)
                );
            }

            [TestMethod, MyFact]
            public void Null()
            {
            myAssert.Throws<InvalidUseOfNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(DBNull.Value);
                });
            }

            [TestMethod, MyFact]
            public void True()
            {
                myAssert.AreEqual(
                    (Int16)(-1),
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(true)
                );
            }

            [TestMethod, MyFact]
            public void False()
            {
                myAssert.AreEqual(
                    (Int16)0,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(false)
                );
            }

            [TestMethod, MyFact]
            public void BlankString()
            {
            myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("");
                });
            }

            [TestMethod, MyFact]
            public void PositiveIntegerString()
            {
                myAssert.AreEqual(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("12")
                );
            }

            [TestMethod, MyFact]
            public void PositiveIntegerStringWithLeadingWhitespace()
            {
                myAssert.AreEqual(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(" 12")
                );
            }

            [TestMethod, MyFact]
            public void PositiveIntegerStringWithTrailingWhitespace()
            {
                myAssert.AreEqual(
                    12d, // VBScript parses string into Doubles, even if there is no decimal fraction
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("12 ")
                );
            }

            [TestMethod, MyFact]
            public void TwoIntegersSeparatedByWhitespace()
            {
            myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("1 1");
                });
            }

            [TestMethod, MyFact]
            public void PositiveDecimalString()
            {
                myAssert.AreEqual(
                    1.2,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("1.2")
                );
            }

            [TestMethod, MyFact]
            public void PseudoNumberWithMultipleDecimalPoints()
            {
                myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("1.1.0");
                });
            }

            [TestMethod, MyFact]
            public void DateAndTime()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                myAssert.AreEqual(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(date)
                );
            }

            [TestMethod, MyFact]
            public void IntegerWithDate()
            {
                myAssert.AreEqual(
                    new DateTime(1899, 12, 31), // This is the VBScript "ZeroDate" plus one day (which is what 1 is translated into in order to become a date)
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(1, new DateTime(2015, 1, 22, 20, 11, 5, 0))
                );
            }

            [TestMethod, MyFact]
            public void BytesWithAnInteger()
            {
                // In a loop "FOR i = CBYTE(1) TO CBYTE(5) STEP 1", the "Integer" step of 1 (this would happen with an implicit step too, since that defaults
                // to an "Integer" 1), the loop variable will be "Integer" since it must be a type that can contain all of the constraints. In order to have
                // a loop variable of type "Byte" the loop would need to be of the form "FOR i = CBYTE(1) TO CBYTE(5) STEP CBYTE(1)".
                myAssert.AreEqual(
                    (Int16)1,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM((byte)1, (byte)5, (Int16)1)
                );
            }

            [TestMethod, MyFact]
            public void DateWithDoublesThatAreWithinDateAcceptableRange()
            {
                var date = new DateTime(2015, 1, 22, 20, 11, 5, 0);
                myAssert.AreEqual(
                    new DateTime(2015, 1, 22, 20, 11, 5, 0),
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(date, 1d)
                );
            }

            [TestMethod, MyFact]
            public void DateWithDoublesThatAreNotWithinDateAcceptableRange()
            {
                myAssert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(new DateTime(2015, 1, 25, 17, 16, 0), double.MaxValue);
                });
            }

            [TestMethod, MyFact]
            public void IntegerWithIntegerValueAsString()
            {
                // Strings are always parsed into doubles, regardless of the size of the value they represent
                myAssert.AreEqual(
                    1d,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM((Int16)1, "2")
                );
            }

            [TestMethod, MyFact]
            public void StringRepresentationsOfDatesAreNotParsed()
            {
            myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("1/1/2015");
                });
            }

            [TestMethod, MyFact]
            public void StringRepresentationsOfISODatesAreNotParsed()
            {
            myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("2015-01-01");
                });
            }

            [TestMethod, MyFact]
            public void StringRepresentationsOfBooleanValuesAreNotParsed()
            {
            myAssert.Throws<TypeMismatchException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM("True");
                });
            }

            [TestMethod, MyFact]
            public void DecimalIsEvenBiggerThanDouble()
            {
                // Although the double type can contain a greater range of values than decimal, VBScript prefers decimal if both are present
                myAssert.AreEqual(
                    1m,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(1m, 2d)
                );
            }

            [TestMethod, MyFact]
            public void DecimalWithDoublesThatAreNotWithinVBScriptCurrencyAcceptableRange()
            {
            // See https://msdn.microsoft.com/en-us/library/9e7a57cf%28v=vs.84%29.aspx for limits of the VBScript data types
            myAssert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(
                        922337203685475m, // Toward the top end of the Currency limit
                        1000000000000000d // Definitely past it
                    );
                });
            }

            [TestMethod, MyFact]
            public void IntegerWithDecimal()
            {
                myAssert.AreEqual(
                    1m,
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NUM(1, 2m)
                );
            }

            // TODO: String with underscores, hashes, exclamation marks

            // TODO: Negatives
            // TODO: Multiple negatives
        }
    //}
}
