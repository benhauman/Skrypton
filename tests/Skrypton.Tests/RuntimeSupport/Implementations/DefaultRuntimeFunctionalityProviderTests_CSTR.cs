
using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.Tests.Shared;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class CSTR
    {
        private readonly CultureInfo TestCulture = CultureInfo.InvariantCulture;

        [TestMethod, MyTheory, MyMemberData("SuccessData")]
        public void SuccessCases(string description, object value, string expectedResult)
        {
            myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR(value));
        }

        [TestMethod, MyTheory, MyMemberData("InvalidUseOfNullData")]
        public void InvalidUseOfNullCases(string description, object value)
        {
            myAssert.Throws<InvalidUseOfNullException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("TypeMismatchData")]
        public void TypeMismatchCases(string description, object value)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("ObjectVariableNotSetData")]
        public void ObjectVariableNotSetCases(string description, object value)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR(value);
            });
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                // Note: CSTR handling of dates varies by culture, so there are tests classes specifically around this further down in this file
                yield return new object[] { "Empty", null, "" };
                yield return new object[] { "Blank string", "", "" };
                yield return new object[] { "Populated string", "abc", "abc" };
                yield return new object[] { "Integer 1", 1, "1" };
                yield return new object[] { "Floating point 1.23", 1.23, "1.23" };
            }
        }

        public static IEnumerable<object[]> InvalidUseOfNullData
        {
            get
            {
                yield return new object[] { "Null", DBNull.Value };
                yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
            }
        }

        public static IEnumerable<object[]> TypeMismatchData
        {
            get
            {
                yield return new object[] { "An empty array", new object[0] };
                yield return new object[] { "Object with default property which is an empty array", new exampledefaultpropertytype { result = new object[0] } };
            }
        }

        public static IEnumerable<object[]> ObjectVariableNotSetData
        {
            get
            {
                yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
            }
        }

        [TestClass]
        public class en_GB : CultureOverridingTests
        {
            public en_GB() : base("en-GB") { }

            [TestMethod, MyTheory, MyMemberData("SuccessData")]
            public void SuccessCases(string description, object value, string expectedResult)
            {
                myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR(value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "28/05/2015" };
                    yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "28/05/2015 18:54:36" };
                    yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "18:54:36" };
                    yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "00:00:00" };
                }
            }
        }

        [TestClass]
        public class en_US : CultureOverridingTests
        {
            public en_US() : base("en-US")
            { }

            [TestMethod, MyTheory, MyMemberData("SuccessData")]
            public void SuccessCases(string description, object value, string expectedResult)
            {
                myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CSTR((DateTime)value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "5/28/2015" }; // new CultureInfo("en-US").DateTimeFormat.ShortDatePattern => d/M/yyyy
                    yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "5/28/2015 6:54:36 PM" };
                    yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "6:54:36 PM" };
                    yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "12:00:00 AM" };
                }
            }
        }
    }
    //}
}
