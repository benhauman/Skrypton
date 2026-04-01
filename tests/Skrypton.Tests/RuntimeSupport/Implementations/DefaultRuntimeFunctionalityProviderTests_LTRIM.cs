
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class LTRIM : TestBase
    {
        [TestMethod]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM(null));
        }

        [TestMethod]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM(DBNull.Value));
        }

        [TestMethod]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("\tValue\t"));
        }

        [TestMethod]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("\nValue\n"));
        }

        [TestMethod]
        public void RemovesMultipleLeadingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("  Value"));
        }

        [TestMethod]
        public void RemovesMultipleLeadingButNotTrailingSpaces()
        {
            myAssert.AreEqual("Value   ", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("  Value   "));
        }
    }
    //}
}
