
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
                // Note: As with the EQ tests, there won't be a lot of cases here around comparing values extracted from object references since that logic is dealt with
                // by the VAL method (and once a non-object-reference value has been obtained, the same logic as illustrated below will be followed)
    public class LT : TestBase
    {
        [TestMethod, MyFact]
        public void EmptyIsNotLessThanEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(null, null)
            );
        }

        [TestMethod, MyFact]
        public void NullComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(DBNull.Value, DBNull.Value)
            );
        }

        /// <summary>
        /// Anything compared to Nothing will error, this is just an example case to illustrate that (if ANYTHING would get a free pass it would be DBNull.Value
        /// but not even it does)
        /// </summary>
        [TestMethod, MyFact]
        public void NullComparedToNothingErrors()
        {
            var nothing = VBScriptConstants.Nothing;
            myAssert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(DBNull.Value, nothing);
                });
        }

        [TestMethod, MyFact]
        public void NothingComparedToNothingErrors()
        {
            var nothing = VBScriptConstants.Nothing;
            myAssert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(nothing, nothing);
                });
        }

        // Empty appears to be treated as zero
        [TestMethod, MyFact]
        public void ZeroIsNotLessThanEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(0, null)
            );
        }
        [TestMethod, MyFact]
        public void EmptyIsNotLessThanZero()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(null, 0)
            );
        }
        [TestMethod, MyFact]
        public void MinusOneIsLessThanEmpty()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(-1, null)
            );
        }
        [TestMethod, MyFact]
        public void EmptyIsNotLessThanMinusOne()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(null, -1)
            );
        }
        [TestMethod, MyFact]
        public void EmptyIsLessThanPlusOne()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(null, 1)
            );
        }
        [TestMethod, MyFact]
        public void PlusOneIsNotLessThanEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().LT(1, null)
            );
        }
    }
    //}
}
