
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
        [TestMethod]
        public void Null()
        {
            myAssert.Throws<ArgumentNullException>(() =>
                {
                    DefaultRuntimeSupportClassFactoryInstance.Get().ARRAY(null);
                });
        }

        [TestMethod]
        public void ZeroElements()
        {
            myAssert.AreEqual(new object[0], DefaultRuntimeSupportClassFactoryInstance.Get().ARRAY());
        }

        [TestMethod]
        public void OneElement()
        {
            myAssert.AreEqual(new object[] { 1 }, DefaultRuntimeSupportClassFactoryInstance.Get().ARRAY(1));
        }

        [TestMethod]
        public void TwoElements()
        {
            myAssert.AreEqual(new object[] { 1, 2 }, DefaultRuntimeSupportClassFactoryInstance.Get().ARRAY(1, 2));
        }
    }
    //}
}
