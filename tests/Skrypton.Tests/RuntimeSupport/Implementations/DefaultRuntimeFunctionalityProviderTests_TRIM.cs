
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
        [TestMethod]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().TRIM(null));
        }

        [TestMethod]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().TRIM(DBNull.Value));
        }

        [TestMethod]
        public void DoesNotRemoveTabs()
        {
            myAssert.AreEqual("\tValue\t", DefaultRuntimeSupportClassFactoryInstance.Get().TRIM("\tValue\t"));
        }

        [TestMethod]
        public void DoesNotRemoveLineReturns()
        {
            myAssert.AreEqual("\nValue\n", DefaultRuntimeSupportClassFactoryInstance.Get().TRIM("\nValue\n"));
        }

        [TestMethod]
        public void RemovesMultipleLeadingAndTrailingSpaces()
        {
            myAssert.AreEqual("Value", DefaultRuntimeSupportClassFactoryInstance.Get().TRIM("  Value   "));
        }
    }
    //}
}
