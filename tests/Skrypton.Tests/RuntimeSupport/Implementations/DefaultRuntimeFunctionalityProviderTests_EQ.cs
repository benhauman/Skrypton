

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
                // Note: There are a class of tests that are not present here - where one or both sides of the comparison are an object reference. In these cases, the EQ
                // implementation on the DefaultRuntimeFunctionalityProvider class pushes these through the VAL method in order to extract a value for comparison (if this
                // fails then a Type Mismatch error is raised). When values are present on both sides, the logic tested here is applied. The tests for the VAL method will
                // cover the logic regarding this, we don't need to duplicate it here. The same goes for arrays - the VAL logic will handle it.
    public class EQ : TestBase
    {
        [TestMethod, MyFact, Owner("Luben Naumov")]
        public void NumericEqualNumericText()
        {
            Assert.IsTrue((bool)DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ((object)-2, (object)"-2"));
        }
        [TestMethod, MyFact, Owner("Luben Naumov")]
        public void NumericTextEqualNumeric()
        {
            Assert.IsTrue((bool)DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ((object)"-2", (object)-2));
        }

        [TestMethod, MyFact]
        public void EmptyEqualsEmpty()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(null, null)
            );
        }

        [TestMethod, MyFact]
        public void NullComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(DBNull.Value, DBNull.Value)
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
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(DBNull.Value, nothing);
                });
        }

        [TestMethod, MyFact]
        public void NothingComparedToNothingErrors()
        {
            var nothing = VBScriptConstants.Nothing;
            myAssert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(nothing, nothing);
                });
        }

        [TestMethod, MyFact]
        public void MinusOneDoesNotEqualEmpty()
        {
            // Non-zero numeric values compared to Empty for equality always return false
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(-1, null)
            );
        }

        [TestMethod, MyFact]
        public void PlusOneDoesNotEqualEmpty()
        {
            // Non-zero numeric values compared to Empty for equality always return false
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(1, null)
            );
        }

        [TestMethod, MyFact]
        public void ZeroEqualsEmpty()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(0, null)
            );
        }

        [TestMethod, MyFact]
        public void MinusOneComparedToNullIsNull()
        {
            // Non-zero numeric values compared to Empty for equality always return false
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(1, DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void PlusOneComparedToNullIsNull()
        {
            // Non-zero numeric values compared to Empty for equality always return false
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(-1, DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void ZeroComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(0, DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void MinusOneEqualsTrue()
        {
            // -1 and True are considered to be the same, as are 0 and False
            // - No other numbers are considered to be equals of booleans (not -1.1, not -2, not 1, not 2)
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(-1, true)
            );
        }

        [TestMethod, MyFact]
        public void ZeroEqualsFalse()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(0, false)
            );
        }

        [TestMethod, MyFact]
        public void MinusOnePointOneDoesNotEqualTrue()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(-1.1, true)
            );
        }

        [TestMethod, MyFact]
        public void MinusOnePointOneDoesNotEqualFalse()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(-1.1, false)
            );
        }

        [TestMethod, MyFact]
        public void PlusOneDoesNotEqualTrue()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(1, true)
            );
        }

        [TestMethod, MyFact]
        public void PlusOneDoesNotEqualFalse()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(1, false)
            );
        }

        [TestMethod, MyFact]
        public void EmptyStringEqualsEmpty()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("", null)
            );
        }

        [TestMethod, MyFact]
        public void EmptyStringComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("", DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void EmptyStringDoesNotEqualsTrue()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("", true)
            );
        }

        [TestMethod, MyFact]
        public void EmptyStringDoesNotEqualsFalse()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("", false)
            );
        }

        [TestMethod, MyFact]
        public void WhiteSpaceStringDoesNotEqualEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(" ", null)
            );
        }

        [TestMethod, MyFact]
        public void WhiteSpaceStringComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(" ", DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void NumericContentStringValueDoesNotEqualNumericValue() // @lubo : this is changed due to CNC mailType compare!
        {
            // Recall that the VBScript expression ("12" = 12) will return true, but if v12String = "12" and v12 = 12 then (v12String = v12) will return
            // false. For cases where string or number literals are present in the comparison, the translator must cast the other side so that they both
            // are consistent but the EQ method does not have to deal with it - so, here, "12" does not equal 12.
            myAssert.AreEqual(
                true,//false, !!!! lubo !!!!
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("12", 12)
            );
        }

        [TestMethod, MyFact]
        public void BooleanContentStringValueDoesNotEqualBooleanValue()
        {
            // See the note in NumericContentStringValueDoesNotEqualNumericValue about literals - the same applies here; while ("True" = True) will return
            // true, if vTrueString = "True" and vTrue = True then (vTrueString = vTrue) return false and it is only this latter case that EQ must deal
            // with, any special handling regarding literals must be dealt with by the translator before getting to EQ.
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ("True", true)
            );
        }

        [TestMethod, MyFact]
        public void TrueDoesNotEqualEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(true, null)
            );
        }

        [TestMethod, MyFact]
        public void FalseEqualsEmpty()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(false, null)
            );
        }

        [TestMethod, MyFact]
        public void TrueComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(true, DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void FalseComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(false, DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void TrueEqualsMinusOne()
        {
            // Dim vTrue, vMinusOne: vTrue = True: vMinusOne = -1: If (vTrue = vMinusOne) Then ' True
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(true, -1)
            );
        }

        [TestMethod, MyFact]
        public void TrueEqualsDoubleMinusOne()
        {
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(true, -1.0d)
            );
        }

        [TestMethod, MyFact]
        public void FalseEqualsZero()
        {
            // Dim vFalse, vZero: vFalse = False: vZero = 0: If (vFalse = vZero) Then ' True
            myAssert.AreEqual(
                true,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(false, 0)
            );
        }

        [TestMethod, MyFact]
        public void DateComparedToNullIsNull()
        {
            myAssert.AreEqual(
                DBNull.Value,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(new DateTime(2015, 1, 19, 22, 52, 0), DBNull.Value)
            );
        }

        [TestMethod, MyFact]
        public void NonZeroDateDoesNotEqualEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(new DateTime(2015, 1, 19, 22, 52, 0), null)
            );
        }

        /// <summary>
        /// The ZeroDate is returned from DateSerial(0, 0, 0) and could feasibly be found to match Empty since the sort-of zero values for booleans,
        /// numbers and strings match Empty. However, this is not the case.
        /// </summary>
        [TestMethod, MyFact]
        public void ZeroDateDoesNotEqualEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(VBScriptConstants.ZeroDate, null)
            );
        }

        /// <summary>
        /// This is explains similar logic to ZeroDateDoesNotEqualEmpty - should the minimum value that VBScript can describe (which is potentially
        /// consider zero to its internals) be found to equal Empty? No, it should not.
        /// </summary>
        [TestMethod, MyFact]
        public void EarliestPossibleDateDoesNotEqualEmpty()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(VBScriptConstants.EarliestPossibleDate, null)
            );
        }

        [TestMethod, MyFact]
        public void ZeroDateDoesNotEqualFalse()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(VBScriptConstants.ZeroDate, false)
            );
        }

        [TestMethod, MyFact]
        public void ZeroDateDoesNotEqualZero()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(VBScriptConstants.ZeroDate, 0)
            );
        }

        [TestMethod, MyFact]
        public void ZeroDateDoesNotEqualEmptyString()
        {
            myAssert.AreEqual(
                false,
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().EQ(VBScriptConstants.ZeroDate, "")
            );
        }
    }
    //}
}
