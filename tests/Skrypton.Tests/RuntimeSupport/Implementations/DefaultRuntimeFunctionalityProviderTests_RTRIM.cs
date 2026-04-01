
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
        [TestMethod]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM(null));
        }

        [TestMethod]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM(DBNull.Value));
        }

        [TestMethod]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM("\tValue\t"));
        }

        [TestMethod]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM("\nValue\n"));
        }

        [TestMethod]
        public void RemovesMultipleTrailingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM("Value   "));
        }

        [TestMethod]
        public void RemovesMultipleTrailingButNotLeadingSpaces()
        {
            myAssert.AreEqual("  Value", DefaultRuntimeSupportClassFactoryInstance.Get().RTRIM("  Value   "));
        }
    }
    //}
}
