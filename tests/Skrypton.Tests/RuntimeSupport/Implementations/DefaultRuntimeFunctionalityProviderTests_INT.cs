

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
    public class INT : TestBase
    {
        [TestMethod, MyMemberData("SuccessData")]
        public void SuccessCases(string description, object value, object expectedResult)
        {
            myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactoryInstance.Get().INT(value));
        }

        [TestMethod, MyMemberData("ObjectVariableNotSetData")]
        public void ObjectVariableNotSetCases(string description, object value)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().INT(value);
            });
        }

        [TestMethod, MyMemberData("TypeMismatchData")]
        public void TypeMismatchCases(string description, object value)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().INT(value);
            });
        }

        [TestMethod, MyMemberData("ObjectDoesNotSupportPropertyOrMemberData")]
        public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object value)
        {
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().INT(value);
            });
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                return
                [
                        new object[] { "Empty", null, (Int16)0 },
                        ["Null", DBNull.Value, DBNull.Value],
                        ["True", true, (Int16)(-1)],
                        ["False", false, (Int16)0],
                        ["Byte", (byte)123, (byte)123],
                        ["Integer", (Int16)123, (Int16)123],
                        ["Long (within Integer range)", (Int32)123, (Int32)123],
                        ["Single (within Integer range)", (Single)123, (Single)123],
                        ["Double (within Integer range)", (Double)123, (Double)123],
                        ["Decimal (within Integer range)", (Decimal)123, (Decimal)123],
                        ["Date (removes time component)", new DateTime(2017, 3, 8, 18, 30, 12, 22), new DateTime(2017, 3, 8)],
                        ["String representing numeric value", "123.45", (double)123],
                        ["String representing numeric value with leading and trailing whitespace", " 123.45 ", (double)123],
                        ["Object with default property which is decimal 123.45", new exampledefaultpropertytype { result = 123.45m }, 123m],

						// A few tests to reinforce that the fraction is removed, it's NOT rounded away from zero or even numbers
						["0.5", (double)0.5, (double)0],
                        ["1.5", (double)1.5, (double)1],
                        ["2.5", (double)2.5, (double)2],
                        ["3.5", (double)3.5, (double)3],

						// These results are surprising, I had expected VBScript to remove the fraction from a number like -0.5 to leave 0 (or from -1.5 to leave -1) but it doesn't!
						["-0.5", (double)(-0.5), (double)(-1)],
                        ["-1.5", (double)(-1.5), (double)(-2)],
                        ["-2.5", (double)(-2.5), (double)(-3)],
                        ["-3.5", (double)(-3.5), (double)(-4)]
                    ];
            }
        }

        public static IEnumerable<object[]> ObjectVariableNotSetData
        {
            get
            {
                return [new object[] { "Nothing", VBScriptConstants.Nothing }];
            }
        }

        public static IEnumerable<object[]> TypeMismatchData
        {
            get
            {
                return
                [
                        new object[] { "Blank String", "" },
                        ["Whitespace", " "],
                        ["String representation of boolean", "True"],
                        ["String representing of numeric value whitespace around decimal point", "123. 45"],
                        ["Unintialised array", new object[0]]
                    ];
            }
        }

        public static IEnumerable<object[]> ObjectDoesNotSupportPropertyOrMemberData
        {
            get
            {
                return [new object[] { "Object without default property", new object() }];
            }
        }
    }
    //}
}
