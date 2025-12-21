
using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class JOIN : TestBase
    {
        [TestMethod, MyTheory, MyMemberData("SuccessData")]
        public void SuccessCases(string description, object value, object delimiter, string expectedResult)
        {
            myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().JOIN(value, delimiter));
        }

        [TestMethod, MyTheory, MyMemberData("InvalidUseOfNullData")]
        public void InvalidUseOfNullCases(string description, object value, object delimiter)
        {
            myAssert.Throws<InvalidUseOfNullException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().JOIN(value, delimiter);
            });
        }

        [TestMethod, MyTheory, MyMemberData("TypeMismatchData")]
        public void TypeMismatchCases(string description, object value, object delimiter)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().JOIN(value, delimiter);
            });
        }

        [TestMethod, MyTheory, MyMemberData("ObjectVariableNotSetData")]
        public void ObjectVariableNotSetCases(string description, object value, object delimiter)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().JOIN(value, delimiter);
            });
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return new object[] { "Uninitialised object array", new object[0], " ", "" };
                yield return new object[] { "Uninitialised object array with Empty delimiter", new object[0], null, "" };

                yield return new object[] { "1D object array of numeric values with comma delimiter", new object[] { 1, 2, 3 }, ",", "1,2,3" };
                yield return new object[] { "1D object array of numeric values with comma+space delimiter", new object[] { 1, 2, 3 }, ", ", "1, 2, 3" };
                yield return new object[] { "1D object array of numeric/Empty values with comma delimiter", new object[] { 1, null, 3 }, ",", "1,,3" };
            }
        }

        public static IEnumerable<object[]> InvalidUseOfNullData
        {
            get
            {
                yield return new object[] { "Null", DBNull.Value, " " };
                yield return new object[] { "Null delimiter", new object[0], DBNull.Value };
                yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, " " };
            }
        }

        public static IEnumerable<object[]> TypeMismatchData
        {
            get
            {
                yield return new object[] { "Empty", null, " " };
                yield return new object[] { "Zero", 0, " " };
                yield return new object[] { "Blank string", "", " " };
                yield return new object[] { "String: \"Test\"", "Test", " " };
                yield return new object[] { "2D object array", new object[0, 0], " " };
                yield return new object[] { "1D object array of numeric/Null values with comma delimiter", new object[] { 1, DBNull.Value, 3 }, "," }; // Would have expected invalid-use-of-null! But VBScript goes for type-mismatch..
                yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" }, " " };
            }
        }

        public static IEnumerable<object[]> ObjectVariableNotSetData
        {
            get
            {
                yield return new object[] { "Nothing", VBScriptConstants.Nothing, " " };
                yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing }, " " };
                yield return new object[] { "1D object array of numeric/Nothing values with comma delimiter", new object[] { 1, VBScriptConstants.Nothing, 3 }, "," };
            }
        }
    }
    //}
}
