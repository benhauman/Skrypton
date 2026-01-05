
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
        [TestMethod, MyFact]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM(null));
        }

        [TestMethod, MyFact]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM(DBNull.Value));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("\tValue\t"));
        }

        [TestMethod, MyFact]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("\nValue\n"));
        }

        [TestMethod, MyFact]
        public void RemovesMultipleLeadingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("  Value"));
        }

        [TestMethod, MyFact]
        public void RemovesMultipleLeadingButNotTrailingSpaces()
        {
            myAssert.AreEqual("Value   ", DefaultRuntimeSupportClassFactoryInstance.Get().LTRIM("  Value   "));
        }
    }
    //}
}
