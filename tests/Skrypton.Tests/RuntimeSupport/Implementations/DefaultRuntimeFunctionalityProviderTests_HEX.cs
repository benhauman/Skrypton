

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
    public class HEX : TestBase
    {
        [TestMethod, MyTheory, MyMemberData("SuccessData")]
        public void SuccessCases(string description, object value, object expectedResult)
        {
            myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactoryInstance.Get().HEX(value));
        }

        [TestMethod, MyTheory, MyMemberData("ObjectVariableNotSetData")]
        public void ObjectVariableNotSetCases(string description, object value)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().HEX(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("TypeMismatchData")]
        public void TypeMismatchCases(string description, object value)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().HEX(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("ObjectDoesNotSupportPropertyOrMemberData")]
        public void ObjectDoesNotSupportPropertyOrMemberCases(string description, object value)
        {
            myAssert.Throws<ObjectDoesNotSupportPropertyOrMemberException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().HEX(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("OverflowData")]
        public void OverflowCases(string description, object value)
        {
            myAssert.Throws<VBScriptOverflowException>(() =>
            {
                DefaultRuntimeSupportClassFactoryInstance.Get().HEX(value);
            });
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                return
                [
						// Unlike some functions, Null IS acceptable
						new object[] { "Null", DBNull.Value, DBNull.Value },

						// Zero-like values
						["Empty", null, "0"],
                        ["0 (Integer)", (short)0, "0"],
                        ["0 (Double)", 0d, "0"],
                        ["False", false, "0"],

						// Larger positive values
						["1 (Byte)", (byte)1, "1"],
                        ["1 (Integer)", (short)1, "1"],
                        ["1 (Currency)", 1m, "1"],
                        ["1 (Single)", 1f, "1"],
                        ["32767 (Integer)", (short)32767, "7FFF"],
                        ["32768 (Long)", 32768, "8000"],
                        ["2147483647 (Long)", 2147483647, "7FFFFFFF"], // Largest positive numer acceptable before overflow

						// -1 values
						["-1 (Integer)", (short)(-1), "FFFF"],
                        ["-2 (Integer)", (short)(-2), "FFFE"],
                        ["-1 (Double)", -1d, "FFFFFFFF"],
                        ["-2 (Double)", -2d, "FFFFFFFE"],
                        ["-1 (String)", "-1", "FFFFFFFF"],
                        ["True", true, "FFFF"],

						// Larger negative values
						["-32767 (Integer)", (short)(-32767), "8001"],
                        ["-32768 (Long)", -32768, "FFFF8000"],
                        ["-2147483648 (Double)", -2147483648d, "80000000"], // Largest negative numer acceptable before overflow

						// A few tests to reinforce that the rounding of numbers works as required
						["0.1 (Double)", 0.1d, "0"],
                        ["0.4 (Double)", 0.4d, "0"],
                        ["0.5 (Double)", 0.5d, "0"],
                        ["0.6 (Double)", 0.6d, "1"],
                        ["1.1 (Double)", 1.1d, "1"],
                        ["1.4 (Double)", 1.4d, "1"],
                        ["1.5 (Double)", 1.5d, "2"],
                        ["1.6 (Double)", 1.6d, "2"],
                        ["2.1 (Double)", 2.1d, "2"],
                        ["2.4 (Double)", 2.4d, "2"],
                        ["2.5 (Double)", 2.5d, "2"],
                        ["2.6 (Double)", 2.6d, "3"],
                        ["3.1 (Double)", 3.1d, "3"],
                        ["3.4 (Double)", 3.4d, "3"],
                        ["3.5 (Double)", 3.5d, "4"],
                        ["3.6 (Double)", 3.6d, "4"],
                        ["-0.1 (Double)", -0.1d, "0"],
                        ["-0.4 (Double)", -0.4d, "0"],
                        ["-0.5 (Double)", -0.5d, "0"],
                        ["-0.6 (Double)", -0.6d, "FFFFFFFF"],
                        ["-1 (Double)", -1d, "FFFFFFFF"],
                        ["-1.1 (Double)", -1.1d, "FFFFFFFF"],
                        ["-1.4 (Double)", -1.4d, "FFFFFFFF"],
                        ["-1.5 (Double)", -1.5d, "FFFFFFFE"],
                        ["-1.6 (Double)", -1.6d, "FFFFFFFE"],
                        ["-2.1 (Double)", -2.1d, "FFFFFFFE"],
                        ["-2.4 (Double)", -2.4d, "FFFFFFFE"],
                        ["-2.5 (Double)", -2.5d, "FFFFFFFE"],
                        ["-2.6 (Double)", -2.6d, "FFFFFFFD"],
                        ["-3.1 (Double)", -3.1d, "FFFFFFFD"],
                        ["-3.4 (Double)", -3.4d, "FFFFFFFD"],
                        ["-3.5 (Double)", -3.5d, "FFFFFFFC"],
                        ["-3.6 (Double)", -3.6d, "FFFFFFFC"]
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

        public static IEnumerable<object[]> OverflowData
        {
            get
            {
                return
                [
                        new object[] { "2147483648", 2147483648 },
                        ["-2147483649", -2147483649]
                    ];
            }
        }
    }
    //}
}
