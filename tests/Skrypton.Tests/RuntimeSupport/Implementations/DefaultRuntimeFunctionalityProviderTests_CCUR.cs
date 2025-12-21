
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class CCUR : TestBase
    {
        [TestMethod, MyFact]
        public void JustBeforePositiveOverflow()
        {
            myAssert.AreEqual(
                VBScriptConstants.MaxCurrencyValue,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CCUR(VBScriptConstants.MaxCurrencyValue)
            );
        }

        [TestMethod, MyFact]
        public void PositiveOverflow()
        {
            myAssert.Throws<VBScriptOverflowException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CCUR(VBScriptConstants.MaxCurrencyValue + 0.000001m);
            });
        }

        [TestMethod, MyFact]
        public void JustBeforeNegativeOverflow()
        {
            myAssert.AreEqual(
                VBScriptConstants.MinCurrencyValue,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CCUR(VBScriptConstants.MinCurrencyValue)
            );
        }

        [TestMethod, MyFact]
        public void NegativeOverflow()
        {
            myAssert.Throws<VBScriptOverflowException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CCUR(VBScriptConstants.MinCurrencyValue - 0.000001m);
                });
        }
    }
    //}
}
