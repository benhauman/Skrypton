

using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class LCASE : TestBase
    {
        [TestMethod]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().LCASE(null));
        }

        [TestMethod]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().LCASE(DBNull.Value));
        }

        [TestMethod]
        public void Test()
        {
            myAssert.AreEqual("test", DefaultRuntimeSupportClassFactoryInstance.Get().LCASE("Test"));
        }
    }
    //}
}
