

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
		public class TIMESERIAL : TestBase
    {
			[TestMethod, MyMemberData("SuccessData")]
			public void SuccessCases(string description, object hours, object minutes, object seconds, DateTime expectedResult)
			{
				myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactoryInstance.Get().TIMESERIAL(hours, minutes, seconds));
			}

			[TestMethod, MyMemberData("TypeMismatchData")]
			public void TypeMismatchCases(string description, object hours, object minutes, object seconds)
			{
				myAssert.Throws<TypeMismatchException>(() =>
				{
					DefaultRuntimeSupportClassFactoryInstance.Get().TIMESERIAL(hours, minutes, seconds);
				});
			}

			[TestMethod, MyMemberData("InvalidUseOfNullData")]
			public void InvalidUseOfNullCases(string description, object hours, object minutes, object seconds)
			{
				myAssert.Throws<InvalidUseOfNullException>(() =>
				{
					DefaultRuntimeSupportClassFactoryInstance.Get().TIMESERIAL(hours, minutes, seconds);
				});
			}

			[TestMethod, MyMemberData("ObjectVariableNotSetData")]
			public void ObjectVariableNotSetCases(string description, object hours, object minutes, object seconds)
			{
				myAssert.Throws<ObjectVariableNotSetException>(() =>
				{
					DefaultRuntimeSupportClassFactoryInstance.Get().TIMESERIAL(hours, minutes, seconds);
				});
			}

			[TestMethod, MyMemberData("OverflowData")]
			public void OverflowCases(string description, object hours, object minutes, object seconds)
			{
				myAssert.Throws<VBScriptOverflowException>(() =>
				{
					DefaultRuntimeSupportClassFactoryInstance.Get().TIMESERIAL(hours, minutes, seconds);
				});
			}

			public static IEnumerable<object[]> SuccessData
			{
				get
				{
					return
                    [
                        new object[] { "2 hour(s), 1 minute(s), 0 second(s)", 2, 1, 0, new DateTime(1899, 12, 30, 2, 1, 0) },
						["-2 hour(s), 1 minute(s), 0 second(s)", -2, 1, 0, new DateTime(1899, 12, 30, 1, 59, 0)],
						["-2 hour(s), 0 minute(s), 1 second(s)", -2, 0, 1, new DateTime(1899, 12, 30, 1, 59, 59)],
						["-2 hour(s), 1 minute(s), 1 second(s)", -2, 1, 1, new DateTime(1899, 12, 30, 1, 58, 59)],
						["-2 hour(s), -1 minute(s), -1 second(s)", -2, -1, -1, new DateTime(1899, 12, 30, 2, 1, 1)],

						["10 hour(s), 1 minute(s), 0 second(s)", 10, 1, 0, new DateTime(1899, 12, 30, 10, 1, 0)],
						["-10 hour(s), 1 minute(s), 0 second(s)", -10, 1, 0, new DateTime(1899, 12, 30, 9, 59, 0)],
						["12 hour(s), 1 minute(s), 0 second(s)", 12, 1, 0, new DateTime(1899, 12, 30, 12, 1, 0)],
						["-12 hour(s), 1 minute(s), 0 second(s)", -12, 1, 0, new DateTime(1899, 12, 30, 11, 59, 0)],
						["14 hour(s), 1 minute(s), 0 second(s)", 14, 1, 0, new DateTime(1899, 12, 30, 14, 1, 0)],
						["-14 hour(s), 1 minute(s), 0 second(s)", -14, 1, 0, new DateTime(1899, 12, 30, 13, 59, 0)],
						["26 hour(s), 1 minute(s), 0 second(s)", 26, 1, 0, new DateTime(1899, 12, 31, 2, 1, 0)],
						["-26 hour(s), 1 minute(s), 0 second(s)", -26, 1, 0, new DateTime(1899, 12, 29, 1, 59, 0)],
						["2 hour(s), 80 minute(s), 0 second(s)", 2, 80, 0, new DateTime(1899, 12, 30, 3, 20, 0)],
						["2 hour(s), -80 minute(s), 0 second(s)", 2, -80, 0, new DateTime(1899, 12, 30, 0, 40, 0)],
						["-2 hour(s), 80 minute(s), 0 second(s)", -2, 80, 0, new DateTime(1899, 12, 30, 0, 40, 0)],
						["-2 hour(s), -80 minute(s), 0 second(s)", -2, -80, 0, new DateTime(1899, 12, 30, 3, 20, 0)],

						["2 hour(s), 8000 minute(s), 0 second(s)", 2, 8000, 0, new DateTime(1900, 1, 4, 15, 20, 0)],
						["2 hour(s), -8000 minute(s), 0 second(s)", 2, -8000, 0, new DateTime(1899, 12, 25, 11, 20, 0)],
						["-2 hour(s), 8000 minute(s), 0 second(s)", -2, 8000, 0, new DateTime(1900, 1, 4, 11, 20, 0)],
						["-2 hour(s), -8000 minute(s), 0 second(s)", -2, -8000, 0, new DateTime(1899, 12, 25, 15, 20, 0)],
						["2 hour(s), 0 minute(s), 8000 second(s)", 2, 0, 8000, new DateTime(1899, 12, 30, 4, 13, 20)],
						["2 hour(s), 0 minute(s), -8000 second(s)", 2, 0, -8000, new DateTime(1899, 12, 30, 0, 13, 20)],
						["-2 hour(s), 0 minute(s), 8000 second(s)", -2, 0, 8000, new DateTime(1899, 12, 30, 0, 13, 20)],
						["-2 hour(s), 0 minute(s), -8000 second(s)", -2, 0, -8000, new DateTime(1899, 12, 30, 4, 13, 20)],

						// CDate(0) is the VBScript zero date plus ten days - the calculation should be reversed to translate a date back into a numeric value
						// (so passing 0, 0, CDate(10) should get the same result as passing 0, 0, 10)
						["0, 0, CDate(0)", 0, 0, VBScriptConstants.ZeroDate.AddDays(10), new DateTime(1899, 12, 30, 0, 0, 10)],

						// Check string parsing
						["0, 0, '10'", 0, 0, "10", new DateTime(1899, 12, 30, 0, 0, 10)],

						// .NET null / VBScript Empty ok (and treated as zero)
						["0, 0, Empty", 0, 0, null, new DateTime(1899, 12, 30, 0, 0, 0)],

						// Check rounding (rounds toward even values, as is common elsewhere)
						["0 hour(s), 0 minute(s), 1.4 second(s)", 0, 0, 1.4, new DateTime(1899, 12, 30, 0, 0, 1)],
						["0 hour(s), 0 minute(s), 1.5 second(s)", 0, 0, 1.5, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), 1.6 second(s)", 0, 0, 1.6, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), 2.4 second(s)", 0, 0, 2.4, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), 2.5 second(s)", 0, 0, 2.5, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), 2.6 second(s)", 0, 0, 2.6, new DateTime(1899, 12, 30, 0, 0, 3)],
						["0 hour(s), 0 minute(s), -1.4 second(s)", 0, 0, -1.4, new DateTime(1899, 12, 30, 0, 0, 1)],
						["0 hour(s), 0 minute(s), -1.5 second(s)", 0, 0, -1.5, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), -1.6 second(s)", 0, 0, -1.6, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), -2.4 second(s)", 0, 0, -2.4, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), -2.5 second(s)", 0, 0, -2.5, new DateTime(1899, 12, 30, 0, 0, 2)],
						["0 hour(s), 0 minute(s), -2.6 second(s)", 0, 0, -2.6, new DateTime(1899, 12, 30, 0, 0, 3)],

						// These test the upper bounds of acceptable values before overflow
						["0 hour(s), 0 minute(s), 32767 second(s)", 0, 0, 32767, new DateTime(1899, 12, 30, 9, 6, 7)],
						["0 hour(s), 32767 minute(s), 0 second(s)", 0, 32767, 0, new DateTime(1900, 1, 21, 18, 7, 0)],
						["32767 hour(s), 0 minute(s), 0 second(s)", 32767, 0, 0, new DateTime(1903, 9, 26, 7, 0, 0)],
						["0 hour(s), 0 minute(s), -32768 second(s)", 0, 0, -32768, new DateTime(1899, 12, 30, 9, 6, 8)],
						["0 hour(s), -32768 minute(s), 0 second(s)", 0, -32768, 0, new DateTime(1899, 12, 8, 18, 8, 0)],
						["-32768 hour(s), 0 minute(s), 0 second(s)", -32768, 0, 0, new DateTime(1896, 4, 4, 8, 0, 0)],
					];
				}
			}

			public static IEnumerable<object[]> TypeMismatchData
			{
				get
				{
					return
                    [
                        new object[] { "0, 0, 'abc'", 0, 0, "abc" }
					];
				}
			}

			public static IEnumerable<object[]> InvalidUseOfNullData
			{
				get
				{
					return
                    [
                        new object[] { "0, 0, Null", 0, 0, DBNull.Value },
						["0, 0, Object-with-default-property-which-is-Null", 0, 0, new exampledefaultpropertytype { result = DBNull.Value }]
					];
				}
			}

			public static IEnumerable<object[]> ObjectVariableNotSetData
			{
				get
				{
					return
                    [
                        new object[] { "0, 0, Nothing", 0, 0, VBScriptConstants.Nothing },
						["0, 0, Object-with-default-property-which-is-Nothing", 0, 0, new exampledefaultpropertytype { result = VBScriptConstants.Nothing }],

						// Arguments are evaluated right-to-left (so the "Nothing" causes the error in "Null, Null, Nothing")
						["Null, Null, Nothing", DBNull.Value, DBNull.Value, VBScriptConstants.Nothing],
						["Null, Nothing, 0", DBNull.Value, VBScriptConstants.Nothing, 0]
					];
				}
			}

			public static IEnumerable<object[]> OverflowData
			{
				get
				{
					return
                    [
                        new object[] { "32768 hours", 32768, 0, 0 },
						["32768 minutes", 0, 32768, 0],
						["32768 seconds", 0, 0, 32768],
						["-32769 hours", -32769, 0, 0],
						["-32769 minutes", 0, -32769, 0],
						["-32769 seconds", 0, 0, -32769],
					];
				}
			}
		}
	//}
}
