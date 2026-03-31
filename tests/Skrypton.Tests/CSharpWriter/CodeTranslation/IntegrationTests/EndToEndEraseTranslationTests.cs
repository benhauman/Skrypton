using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.StageTwoParser.ExpressionParsing;
using static Skrypton.Tests.RuntimeSupport.Implementations.VBScriptEsqueValueRetrieverTests;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.Shared.Comparers;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndEraseTranslationTests : TestBase
    {
        [TestMethod, MyTheory, MyMemberData(nameof(SuccessData))]
        public void SuccessCases(int testno, string description, string source, string expected)
        {
            TestCSharpCodeTranslationWithoutScaffolding(expected, source, ["SKY101"]);
        }
        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return [1, "Empty ERASE is a runtime error", "ERASE", "throw new InvalidOperationException(\"Wrong number of arguments: 'Erase' (line 1)\");"];
                yield return [2, "Empty ERASE is a runtime error (with CALL keyword)", "CALL ERASE", "throw new InvalidOperationException(\"Wrong number of arguments: 'Erase' (line 1)\");"];

                yield return [3, "Simplest case: ERASE a", "ERASE a", "_.ERASE(_env.a, v => { _env.a = v; });"];
                yield return [4, "Simplest case: ERASE a (with CALL keyword)", "CALL ERASE(a)", "_.ERASE(_env.a, v => { _env.a = v; });"];

                // If the target is specified with arguments, then it must be an array where the arguments are indices. The non-by-ref ERASE method signature is used and validation of the
                // target (whether it's an array and whether the indices are valid) is handled at runtime.
                yield return [5, "Target with arguments: ERASE a(0)", "ERASE a(0)", "_.ERASE(_env.a, (Int16)0);"];
                yield return [6, "Target with arguments: CALL ERASE(a(0)) (with CALL keyword)", "CALL ERASE(a(0))", "_.ERASE(_env.a, (Int16)0);"];

                // "ERASE a()" is either a "Subscript out of range" or a "Type mismatch", depending upon whether "a" is an array or not - this needs to be decided at runtime. It does this
                // using the non-by-ref argument argument signature. This is the case where "a" is known to be a variable (whether explicitly declared or not, if "a" is known to be a
                // function then it's a different error case).
                yield return [7, "ERASE a()", "ERASE a()", "_.ERASE(_env.a);"];

                yield return
                [
                    8,
                    "Error if the target is known not to be a variable",
                    "ERASE a\nFUNCTION a\nEND FUNCTION",
                        @"var invalidEraseTarget = _.CALLm1v0(this, _outer, ""a"");
                        throw new TypeMismatchException(""'Erase' (line 1)"");
                        public object a()
                        {
                        return null;
                        }"
                ];
                yield return
                [
                    9,
                    "Error if the target is known not to be a variable (takes precedence over other ERASE a() error case)",
                    "ERASE a()\nFUNCTION a\nEND FUNCTION",
                        @"var invalidEraseTarget = _.CALLm1argp(this, _outer, ""a"", _.ARGS.ForceBrackets());
                        throw new TypeMismatchException(""'Erase' (line 1)"");
                        public object a()
                        {
                            return null;
                        }"
                ];

                // Note: When the arguments are invalid, they are still evaluated and THEN the runtime error is raised. The references are not forced into value types (if they appear valid
                // at this point then the ERASE call must confirm at runtime that the target is an array), so the evaulation of some targets (eg. "a") will have no effect while others (eg.
                // "a.GetName()" may have side effects).
                yield return
                [
                    10,
                    "Brackets around target (would be by-val => invalid)",
                    "ERASE (a)",
                        @"var invalidEraseTarget = _env.a;
                        throw new TypeMismatchException(""'Erase' (line 1)"");"
                ];
                yield return
                [
                    11,
                    "Multiple targets",
                    "ERASE a, b",
                        @"var invalidEraseTarget = _env.a;
                        var invalidEraseTarget2 = _env.b;
                        throw new InvalidOperationException(""Wrong number of arguments: 'Erase' (line 1)"");"
                ];
                yield return
                [
                    12,
                    "Member access target",
                    "ERASE a.Name",
                        @"var invalidEraseTarget = _.CALLm1v0(this, _env.a, ""Name"");
                        throw new TypeMismatchException(""'Erase' (line 1)"");"
                ];
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
                public object F1(ref object a)
                {
                    object F1_retVal = null;
                    object byrefalias = a;
                    try
                    {
                        _.ERASE(byrefalias, v => { byrefalias = v; });
                    }
                    finally { a = byrefalias; }
                    return F1_retVal;
                }";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
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
        public static void AreEqualString(string expected, string actual)
        {
            Assert.AreEqual(expected, actual);
        }
        public static void AreEqual<T>(T expected, T actual) // use 'TestCSharpCodeTranslationWithoutScaffolding'
        {
            AreEqualCore<T>(expected, actual);
        }
        internal static int FindArrayStringDiff(string[] arr_e, string[] arr_a)
        {
            if (arr_e.Length >= arr_a.Length)
            {
                for (int idx = 0; idx < arr_e.Length; idx++)
                {
                    string item_a = arr_a.Length <= idx ? null
                        : arr_a[idx];
                    string item_e = arr_e[idx];
                    if (!string.Equals(item_e, item_a, StringComparison.Ordinal))
                        return idx;
                }
            }
            else
            {
                for (int idx = 0; idx < arr_a.Length; idx++)
                {
                    string item_e = arr_e.Length <= idx ? null
                        : arr_e[idx];
                    string item_a = arr_a[idx];
                    if (!string.Equals(item_e, item_a, StringComparison.Ordinal))
                        return idx;
                }
            }
            return -1;
        }
        private static void AreEqualCore<T>(T expected, T actual) // use 'TestCSharpCodeTranslationWithoutScaffolding'
        {
            {
                if (expected is string[] arr_e)
                {
                    string[] arr_a = actual as string[];
                    int idx = FindArrayStringDiff(arr_e, arr_a);
                    if (idx >= 0)
                    {
                        string item_a = arr_a.Length <= idx ? null
                            : arr_a[idx];
                        string item_e = arr_e.Length <= idx ? null
                            : arr_e[idx];
                        Assert.AreEqual(item_e, item_a, message: $"index:{idx}");
                    }
                    return;
                }
            }
            {
                IEnumerable<object> arr_obj_e = expected as IEnumerable<object>;
                if (arr_obj_e != null)
                {
                    myAssert.AreEqualU<T>(expected, actual, myAssert.GetEqualityComparer<T>(null));
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
                if (expected is string str_e && actual is string str_a)
                {
                    Assert.AreEqual(str_e, str_a);
                    return;
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
        public static void AreEqualU<T>(T expected, T actual, IEqualityComparer<T> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }
        public static void AreEqual(PseudoField expected, object actual, IEqualityComparer<object> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }
        public static void AreEqual(IEnumerable<ParsingExpression> expected, IEnumerable<ParsingExpression> actual, IEqualityComparer<IEnumerable<ParsingExpression>> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }
        public static void AreEqual(IEnumerable<ICodeBlock> expected, IEnumerable<ICodeBlock> actual, IEqualityComparer<IEnumerable<ICodeBlock>> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }
        public static void AreEqual(IEnumerable<IToken> expected, IEnumerable<IToken> actual, IEqualityComparer<IEnumerable<IToken>> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }

        public static void AreEqualX(string expectedTranslatedContent, IReadOnlyCollection<NameToken> expectedVariablesAccessed, TranslatedStatementContentDetails actual)
        {
            if (!TranslatedStatementContentDetailsComparer.EqualsX(expectedTranslatedContent, expectedVariablesAccessed, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expectedTranslatedContent + ", Actual:" + actual.TranslatedContent);
            }
        }
        public static void AreEqualCollection<T>(IEnumerable<T> expected, IEnumerable<T> actual, IEqualityComparer<IEnumerable<T>> comparer)
        {
            if (!comparer.Equals(expected, actual))
            {
                Assert.Fail("Not Equal. Expected:" + expected + ", Actual:" + actual);
            }
        }

        internal static void False(bool v, string message = "")
        {
            Assert.IsFalse(v, message);
        }

        internal static void True(bool v, string message = "")
        {
            Assert.IsTrue(v, message);
        }

        internal static void IsNull(object v)
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