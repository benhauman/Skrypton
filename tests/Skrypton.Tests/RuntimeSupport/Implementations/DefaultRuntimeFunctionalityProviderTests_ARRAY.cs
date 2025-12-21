
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class ARRAY : TestBase
    {
        /// <summary>
        /// The ARRAY method should never be called with a null values array - if it is called with zero arguments then the array should be a zero-element array instance, not null
        /// </summary>
        [TestMethod, MyFact]
        public void Null()
        {
            myAssert.Throws<ArgumentNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ARRAY(null);
                });
        }

        [TestMethod, MyFact]
        public void ZeroElements()
        {
            myAssert.AreEqual(new object[0], DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ARRAY());
        }

        [TestMethod, MyFact]
        public void OneElement()
        {
            myAssert.AreEqual(new object[] { 1 }, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ARRAY(1));
        }

        [TestMethod, MyFact]
        public void TwoElements()
        {
            myAssert.AreEqual(new object[] { 1, 2 }, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ARRAY(1, 2));
        }
    }
    //}
}
