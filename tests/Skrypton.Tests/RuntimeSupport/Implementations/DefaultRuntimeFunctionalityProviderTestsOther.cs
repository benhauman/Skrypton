
using Skrypton.RuntimeSupport;
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class NullableNUM : TestBase
    {
        [TestMethod, MyFact]
        public void NullToNumber()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().NullableNUM(DBNull.Value)
            );
        }
    }
    //}
}
