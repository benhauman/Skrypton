

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
        [TestMethod, MyFact]
        public void EmptyResultsInBlankString()
        {
            myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LCASE(null));
        }

        [TestMethod, MyFact]
        public void NullResultsInNull()
        {
            myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LCASE(DBNull.Value));
        }

        [TestMethod, MyFact]
        public void Test()
        {
            myAssert.AreEqual("test", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LCASE("Test"));
        }
    }
    //}
}
