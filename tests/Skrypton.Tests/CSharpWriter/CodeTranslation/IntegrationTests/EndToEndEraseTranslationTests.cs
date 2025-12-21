using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndEraseTranslationTests : TestBase
    {
        [TestMethod, MyTheory, MyMemberData("SuccessData")]
        public void SuccessCases(string description, string source, string[] expected)
        {
            myAssert.AreEqual(expected, WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies));
        }

        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return new object[] { "Empty ERASE is a runtime error", "ERASE", new[] { "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");" } };
                yield return new object[] { "Empty ERASE is a runtime error (with CALL keyword)", "CALL ERASE", new[] { "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");" } };

                yield return new object[] { "Simplest case: ERASE a", "ERASE a", new[] { "_.ERASE(_env.a, v1 => { _env.a = v1; });" } };
                yield return new object[] { "Simplest case: ERASE a (with CALL keyword)", "CALL ERASE(a)", new[] { "_.ERASE(_env.a, v1 => { _env.a = v1; });" } };

                // If the target is specified with arguments, then it must be an array where the arguments are indices. The non-by-ref ERASE method signature is used and validation of the
                // target (whether it's an array and whether the indices are valid) is handled at runtime.
                yield return new object[] { "Target with arguments: ERASE a(0)", "ERASE a(0)", new[] { "_.ERASE(_env.a, (Int16)0);" } };
                yield return new object[] { "Target with arguments: CALL ERASE(a(0)) (with CALL keyword)", "CALL ERASE(a(0))", new[] { "_.ERASE(_env.a, (Int16)0);" } };

                // "ERASE a()" is either a "Subscript out of range" or a "Type mismatch", depending upon whether "a" is an array or not - this needs to be decided at runtime. It does this
                // using the non-by-ref argument argument signature. This is the case where "a" is known to be a variable (whether explicitly declared or not, if "a" is known to be a
                // function then it's a different error case).
                yield return new object[] { "ERASE a()", "ERASE a()", new[] { "_.ERASE(_env.a);" } };

                yield return new object[] {
                    "Error if the target is known not to be a variable",
                    "ERASE a\nFUNCTION a\nEND FUNCTION",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(this, _outer, \"a\");",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");",
                        "public object a()",
                        "{",
                        "return null;",
                        "}"
                    }
                };
                yield return new object[] {
                    "Error if the target is known not to be a variable (takes precedence over other ERASE a() error case)",
                    "ERASE a()\nFUNCTION a\nEND FUNCTION",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(this, _outer, \"a\", _.ARGS.ForceBrackets());",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");",
                        "public object a()",
                        "{",
                        "return null;",
                        "}"
                    }
                };

                // Note: When the arguments are invalid, they are still evaluated and THEN the runtime error is raised. The references are not forced into value types (if they appear valid
                // at this point then the ERASE call must confirm at runtime that the target is an array), so the evaulation of some targets (eg. "a") will have no effect while others (eg.
                // "a.GetName()" may have side effects).
                yield return new object[] {
                    "Brackets around target (would be by-val => invalid)",
                    "ERASE (a)",
                    new[] {
                        "var invalidEraseTarget1 = _env.a;",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");"
                    }
                };
                yield return new object[] {
                    "Multiple targets",
                    "ERASE a, b",
                    new[] {
                        "var invalidEraseTarget1 = _env.a;",
                        "var invalidEraseTarget2 = _env.b;",
                        "throw new Exception(\"Wrong number of arguments: 'Erase' (line 1)\");"
                    }
                };
                yield return new object[] {
                    "Member access target",
                    "ERASE a.Name",
                    new[] {
                        "var invalidEraseTarget1 = _.CALL(this, _env.a, \"Name\");",
                        "throw new TypeMismatchException(\"'Erase' (line 1)\");"
                    }
                };
            }
        }

        [TestMethod, MyFact]
        public void SingleTokenEraseTargetsRequireByRefAliasingIfTheTargetIsByRefArgumentOfTheContainingFunction()
        {
            var source = @"
                Function F1(a)
                    ERASE a
                End Function";
            var expected = @"
                public object f1(ref object a)
                {
                    object retVal1 = null;
                    object byrefalias2 = a;
                    try
                    {
                        _.ERASE(byrefalias2, v3 => { byrefalias2 = v3; });
                    }
                    finally { a = byrefalias2; }
                    return retVal1;
                }";
            myAssert.AreEqual(
                expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Skip(1).Select(v => v.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}

namespace Skrypton.Tests
{
    internal static class myAssert
    {
        public static void ThrowsX(Type exceptionType, Action testCode)
        {
            try
            {
                testCode();
            }
            catch (Exception ex)
            {
                if (ex.GetType() != exceptionType)
                    throw;
            }
        }
        public static void Throws<T>(Action testCode) where T : Exception
        {
            try
            {
                testCode();
            }
            catch (Exception ex)
            {
                if (ex.GetType() != typeof(T))
                    throw;
            }
        }
        public static T Throws<T>(Func<object> testCode) where T : Exception
        {
            try
            {
                return (T)testCode();
            }
            catch (Exception ex)
            {
                if (ex.GetType() != typeof(T))
                    throw;
                return default(T);
            }
        }
        public static void AreEqualDateTime(string msg, DateTime expected, DateTime actual)
        {
            Assert.AreEqual<DateTime>(expected, actual, myAssert.GetEqualityComparer<DateTime>(null), msg);
        }
        public static void AreEqualStringArray(string[] arr_expected, string[] arr_actual)
        {
            string text_e = arr_expected == null ? null : string.Join("\r\n", arr_expected);
            string text_a = arr_actual == null ? null : string.Join("\r\n", arr_actual);
            if (arr_expected != null)
            {
                Assert.IsNotNull(arr_actual, nameof(arr_actual));
                for (int idx = 0; idx < arr_actual.Length; idx++)
                {
                    Assert.AreEqual(arr_actual[idx], arr_actual[idx]);
                }
                return;
            }
            else
            {
                Assert.IsTrue(arr_actual == null || arr_actual.Length == 0);
            }
        }
        public static void AreEqualString(string expected, string actual)
        {
            Assert.AreEqual(expected, actual);
        }
        public static void AreEqual<T>(T expected, T actual)
        {
            {
                string[] arr_e = expected as string[];
                if (arr_e != null)
                {
                    string[] arr_a = actual as string[];
                    for (int idx = 0; idx < arr_e.Length; idx++)
                    {
                        Assert.AreEqual(arr_e[idx], arr_a[idx]);
                    }
                    return;
                }
            }
            {
                IEnumerable<object> arr_obj_e = expected as IEnumerable<object>;
                if (arr_obj_e != null)
                {
                    myAssert.Equal<T>(expected, actual, myAssert.GetEqualityComparer<T>(null));
                    //Xunit.Assert.Equal(expected, actual);
                    ///IEnumerable<object> arr_obj_a = actual as IEnumerable<object>;
                    ///for (int idx = 0; idx < arr_obj_e.Count(); idx++)
                    ///{
                    ///    AreEqual(arr_obj_e.ElementAt(idx), arr_obj_a.ElementAt(idx));
                    ///}
                    return;
                }
            }
            {
                object[] arr_obj_e = expected as object[];
                if (arr_obj_e != null)
                {
                    object[] arr_obj_a = actual as object[];
                    for (int idx = 0; idx < arr_obj_e.Length; idx++)
                    {
                        AreEqual(arr_obj_e[idx], arr_obj_a[idx]);
                    }
                    return;
                }
            }
            {
                double[] arr_obj_e = expected as double[];
                if (arr_obj_e != null)
                {
                    double[] arr_obj_a = actual as double[];
                    for (int idx = 0; idx < arr_obj_e.Length; idx++)
                    {
                        AreEqual(arr_obj_e[idx], arr_obj_a[idx]);
                    }
                    return;
                }
            }
            {
                Single[] arr_obj_e = expected as Single[];
                if (arr_obj_e != null)
                {
                    Single[] arr_obj_a = actual as Single[];
                    for (int idx = 0; idx < arr_obj_e.Length; idx++)
                    {
                        AreEqual(arr_obj_e[idx], arr_obj_a[idx]);
                    }
                    return;
                }
            }
            {
                //if ((object)expected is DateTime dt_e)
                //{
                //    DateTime dt_a = (DateTime)(object)actual;
                //    AreEqual<DateTime>(dt_e, dt_a);
                //    return;
                //}
            }
            {
                if (expected != null && actual != null && expected.GetType() != actual.GetType())
                {
                    if (expected is IConvertible && actual is IConvertible)
                    {
                        var convertedActual = Convert.ChangeType(actual, expected.GetType(), CultureInfo.InvariantCulture);
                        Assert.AreEqual(expected, convertedActual);
                        return;
                    }
                    Assert.AreEqual(expected, actual);
                }
            }
            {
                Assert.AreEqual(expected, actual);
            }
        }
        private static IEqualityComparer<T> GetEqualityComparer<T>(System.Collections.IEqualityComparer innerComparer = null)
        {
            return new AssertEqualityComparer<T>(innerComparer);
        }
        public static void AreEqual<T>(T expected, T actual, IEqualityComparer<T> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }
        public static void Equal<T>(T expected, T actual, IEqualityComparer<T> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }

        internal static void False(bool v)
        {
            Assert.IsFalse(v);
        }

        internal static void True(bool v)
        {
            Assert.IsTrue(v);
        }

        internal static void Null(object v)
        {
            Assert.IsNull(v);
        }

        internal static void IsType<T>(object value)
        {
            Assert.IsInstanceOfType(value, typeof(T));
        }

        internal static void AreEqualX(double expected, double actual, int precision)
        {
            double numE = Math.Round(expected, precision);
            double numA = Math.Round(actual, precision);
            Assert.AreEqual(numE, numA);
        }

        internal static void NotEqual(int expected, int actual)
        {
            Assert.AreNotEqual(expected, actual);
        }
        internal static void NotEqual(float expected, float actual)
        {
            Assert.AreNotEqual(expected, actual);
        }
    }
    sealed class MyFactAttribute : Attribute
    {

    }

    sealed class MyTheoryAttribute : Attribute
    {

    }
    sealed class MyMemberData : Attribute, ITestDataSource //ms.DataRowAttribute // DynamicData
    {
        private string context;
        public MyMemberData(string context)
        {
            this.context = context;
        }

        public IEnumerable<object[]> GetData(MethodInfo methodInfo)
        {
            var pi = methodInfo.DeclaringType.GetProperty(context);
            if (pi == null)
                throw new InvalidOperationException("Property not found:" + context);
            var propertyValue = pi.GetValue(null);
            return (IEnumerable<object[]>)propertyValue;
        }
        public string DisplayName
        {
            get;
            set;
        }

        public string GetDisplayName(MethodInfo methodInfo, object[] data)
        {
            if (!string.IsNullOrWhiteSpace(this.DisplayName))
            {
                return this.DisplayName;
            }

            if (data != null)
            {
                return methodInfo.Name + ":" + string.Join(", ", data.Select(x => "" + x).ToArray());
            }
            else
            {
                return methodInfo.Name + "#" + "<null>";
            }
        }
    }

}