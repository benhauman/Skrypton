
using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.Tests.Shared;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class LEN : TestBase
    {
        [TestMethod, MyFact]
        public void EmptyResultsInZero()
        {
            myAssert.AreEqual((int)0, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LEN(null)); // This should return an int ("Long" in VBScript parlance)
        }

        [TestMethod, MyFact]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LEN(DBNull.Value));
        }

        [TestMethod, MyFact]
        public void Test()
        {
            myAssert.AreEqual(4, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LEN("Test"));
        }

        [TestMethod, MyFact]
        public void NumericValue()
        {
            myAssert.AreEqual(1, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LEN(4)); // Numbers get cast as strings, so the number 4 becomes the string "4" and so has length 1
        }

        [TestClass]
        public class en_GB : CultureOverridingTests
        {
            public en_GB() : base("en-GB") { }

            [TestMethod, MyTheory, MyMemberData("SuccessData")]
            public void SuccessCases(string description, object value, int expectedResult)
            {
                myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LEN(value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Date with zero time", new DateTime(2015, 5, 28), "28/05/2015".Length };
                    yield return new object[] { "Date with non-zero time", new DateTime(2015, 5, 28, 18, 54, 36), "28/05/2015 18:54:36".Length };
                    yield return new object[] { "Zero date with non-zero time", VBScriptConstants.ZeroDate.Add(new TimeSpan(18, 54, 36)), "18:54:36".Length };
                    yield return new object[] { "Zero date with zero time", VBScriptConstants.ZeroDate, "00:00:00".Length };
                }
            }
        }
    }
    //}
}
