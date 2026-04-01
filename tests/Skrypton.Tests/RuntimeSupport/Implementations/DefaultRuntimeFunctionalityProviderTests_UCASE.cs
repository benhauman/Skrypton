
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
    //{
        public class UCASE : TestBase
    {
            [TestMethod]
            public void EmptyResultsInBlankString()
            {
                myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().UCASE(null));
            }

            [TestMethod]
            public void NullResultsInNull()
            {
                myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().UCASE(DBNull.Value));
            }

            [TestMethod]
            public void Test()
            {
                myAssert.AreEqual("TEST", DefaultRuntimeSupportClassFactoryInstance.Get().UCASE("Test"));
            }
        }
    //}
}
