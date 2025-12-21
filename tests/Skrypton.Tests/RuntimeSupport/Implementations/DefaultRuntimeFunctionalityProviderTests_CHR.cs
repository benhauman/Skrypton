

using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class CHR : TestBase
    {
        [TestMethod, MyTheory, MyMemberData(nameof(SuccessData))]
        public void SuccessCases(string description, object value, Char expectedResult)
        {
            myAssert.AreEqual(new string(expectedResult, 1), DefaultRuntimeSupportClassFactoryInstance.Get().CHR(value));
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return new object[] { "Empty", null, (char)0 };
                yield return new object[] { "0", 0, (char)0 };
                yield return new object[] { "1", 1, (char)1 };
                yield return new object[] { "255", 255, (char)255 };
                yield return new object[] { "255.4", 255.4, (char)255 };
                yield return new object[] { "-0.5", -0.5, (char)0 };
                yield return new object[] { "125n", 125, (char)125 };
                yield return new object[] { "125c", 125, '}' };

                yield return new object[] { "233n", 233, (char)233 };
                yield return new object[] { "233c", 233, 'é' };

                yield return new object[] { "8364n", 8364, (char)8364 };
                yield return new object[] { "8364c", 8364, '€' };

                yield return new object[] { "8212n", 8212, (char)8212 };
                yield return new object[] { "8212c", 8212, '—' };

            }
        }

        [TestMethod, MyTheory, MyMemberData(nameof(InvalidUseOfNullData))]
        public void InvalidUseOfNullCases(string description, object value)
        {
            myAssert.Throws<InvalidUseOfNullException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CHR(value);
            });
        }
        public static IEnumerable<object[]> InvalidUseOfNullData
        {
            get
            {
                yield return new object[] { "DBNull", DBNull.Value };
                yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value } };
                yield return new object[] { "null", null };
            }
        }

        [TestMethod, MyTheory, MyMemberData(nameof(TypeMismatchData))]
        public void TypeMismatchCases(string description, object value)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CHR(value);
            });
        }
        public static IEnumerable<object[]> TypeMismatchData
        {
            get
            {
                yield return new object[] { "Blank string", "" };
                yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" } };
            }
        }

        [TestMethod, MyTheory, MyMemberData(nameof(ObjectVariableNotSetData))]
        public void ObjectVariableNotSetCases(string description, object value)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CHR(value);
                });
        }
        public static IEnumerable<object[]> ObjectVariableNotSetData
        {
            get
            {
                yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
            }
        }

        [TestMethod, MyTheory, MyMemberData("InvalidProcedureCallOrArgumentData")]
        public void InvalidProcedureCallOrArgumentCases(string description, object value)
        {
            myAssert.Throws<InvalidProcedureCallOrArgumentException>(() =>
                {
                    DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().CHR(value);
                });
        }

        public static IEnumerable<object[]> InvalidProcedureCallOrArgumentData
        {
            get
            {
                yield return new object[] { "255.5", 255.5 };
                yield return new object[] { "-0.6", -0.6 };
            }
        }
    }
    //}
}
