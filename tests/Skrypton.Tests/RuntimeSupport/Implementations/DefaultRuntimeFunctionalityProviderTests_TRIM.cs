
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class TRIM : TestBase
    {
        [TestMethod, MyFact]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().TRIM(null));
        }

        [TestMethod, MyFact]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().TRIM(DBNull.Value));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().TRIM("\tValue\t"));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().TRIM("\nValue\n"));
        }

        [TestMethod, MyFact]
        public void RemovesMultipleLeadingAndTrailingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().TRIM("  Value   "));
        }
    }
    //}
}
