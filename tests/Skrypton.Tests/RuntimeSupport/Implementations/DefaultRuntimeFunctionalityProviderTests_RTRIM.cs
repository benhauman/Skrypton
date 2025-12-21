
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class RTRIM : TestBase
    {
        [TestMethod, MyFact]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM(null));
        }

        [TestMethod, MyFact]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM(DBNull.Value));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM("\tValue\t"));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM("\nValue\n"));
        }

        [TestMethod, MyFact]
        public void RemovesMultipleTrailingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM("Value   "));
        }

        [TestMethod, MyFact]
        public void RemovesMultipleTrailingButNotLeadingSpaces()
        {
            myAssert.AreEqual("  Value", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().RTRIM("  Value   "));
        }
    }
    //}
}
