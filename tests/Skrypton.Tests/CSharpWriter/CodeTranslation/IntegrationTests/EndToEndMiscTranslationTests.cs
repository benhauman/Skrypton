using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndMiscTranslationTests : TestBase
    {
        // TODO: Test function call with numeric values (1 and 1.1), string values, built-in values and built-in functions (such as "Now") and ensure that
        // they all have the arguments specified as ByVal
        // - Is it easiest to put it in here or better to put it into StatementTranslatorTests?

        /// <summary>
        /// The code here accesses an undeclared variable in a statement in the outermost scope, that scope should be registered in the EnvironmentReferences
        /// class. There is also a "wscript" reference which is declared as an External Dependency in the translator, this will appear in the Environment
        /// References class as well (as any/all External Dependencies should).
        /// </summary>
        [TestMethod]
        public void UndeclaredVariablesInTheOutermostScopeShouldBeDefinedAsAnEnvironmentVariable()
        {
            var source = @"
				WScript.Echo i
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// This code will access an undeclared variable within a function. The scope of that undeclared variable should be restricted to the function in
        /// which it is accessed and not bleed out into the outer scope.
        /// </summary>
        [TestMethod]
        public void UndeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
				Test1
				Function Test1()
					WScript.Echo i
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY103"]);
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [TestMethod]
        public void DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction()
        {
            var source = @"
				Test1
				Function Test1()
					Dim i
					WScript.Echo i
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This is a corresponding test to DeclaredVariableWithinFunctionsShouldBeRestrictedInScopeToThatFunction but for the case where the variable is
        /// explicitly declared.
        /// </summary>
        [TestMethod]
        public void DeclaredVariableInOutermostScopeShouldBeAccessedFromThereWhenRequiredWithinFunction()
        {
            var source = @"
				Dim i
				Test1
				Function Test1()
					WScript.Echo i
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void NumericLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func 1()";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        [TestMethod]
        public void StringLiteralsAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func \"1\"()";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101", "SKY103"]);
        }

        [TestMethod]
        public void BuiltinValuesAccessedAsFunctionsResultInRuntimeErrors()
        {
            var source = "func vbObjectError()";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        [TestMethod]
        public void ClassNameFollowedByBracketsInNewStatementResultsInCompileTimeError()
        {
            var source = "c = new C1()";
            myAssert.Throws<InvalidOperationException>(() =>
            {
                TestCSharpCodeTranslationWithoutScaffolding(source);
            });
        }

        /// <summary>
        /// Since runs of string concatenations are so common, an exception to the two-arguments-per-operation (apart from NOT, that only takes one) is made
        /// to allow the values to be combined in a single CONCAT call, reducing the size of the emitted code
        /// </summary>
        [TestMethod]
        public void ConcatFunctionAllowsMoreThanTwoArguments()
        {
            var source = @"
				WScript.Echo a & b & c & d
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// This is related to the ConcatFunctionAllowsMoreThanTwoArguments and provides reassurance that string concatenations will only be joined if it
        /// would have no effect on the rest of processing (since the addition operation should take precedence, there is no CONCAT-flattening that can
        /// be performed in this case)
        /// </summary>
        [TestMethod]
        public void ConcatFunctionAllowsMoreThanTwoArgumentsButDoesNotAffectNestedOperationsOfOtherTypes()
        {
            var source = @"
				WScript.Echo a & 1 + 2 & c & d
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// The string values that specify target member names in a CALL codeExpression must not be manipulated by the name rewriter at runtime. This means that
        /// their casing will not be affected and - more importantly, any manipulations relating to C# keywords will NOT be applied. When the target is a
        /// translated class, the name rewriter manipulations would not cause any issue but if the target is not something that is translated (a COM component,
        /// for example), then trying to access its members with the name-rewritten versions will fail. This means that the CALL implementation must be able to
        /// consider the same name rewriter rules at runtime that the translator does.
        /// </summary>
        [TestMethod]
        public void MemberAccessorsInCallStatementsShouldNotBeRenamedAtTranslationTime()
        {
            // "Params" is a C# keyword, so we couldn't emit translated code with a method called "Params", but if "a" is an external reference (such as a COM
            // component) then It may have a methor or property named "Params". As such we mustn't enforce the rewriting of "Params" to something C#-friendly
            // at compile time (the CALL implementation will have to do some magic)
            // - The GetTranslatedStatements uses the DefaultTranslator which uses the DefaultRuntimeSupportClassFactory.DefaultNameRewriter which will
            //   ensure that C# keywords are rewritten to something safe
            var source = @"
				WScript.Echo a.Params
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// Similar to MemberAccessorsInCallStatementsShouldNotBeRenamedAtTranslationTime, the ValueSettingStatementsTranslator has been corrected so that it
        /// won't rewrite member accessors that string arguments in a SET call
        /// </summary>
        [TestMethod]
        public void MemberAccessorsInValueSettingsStatementsShouldNotBeRenamedAtTranslationTime()
        {
            var source = @"
				a.Name.Length = 1
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// It doesn't matter if we're within a VBScript class on in the outermost scope, or within a function in the outermost scope, the "Me" reference may
        /// always be mapped directly to "this" and it will be correct
        /// </summary>
        [TestMethod]
        public void MeReferenceMapsDirectlyOnToThis()
        {
            var source = @"
				WScript.Echo Me.Name
			";
            TestCSharpCodeTranslationWithoutScaffolding(source); TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// If a CALL codeExpression has a function as its target then it needs to be rewritten so that that the owner of the function (or property) is the target
        /// and the function and one of the member accessors (since it's not valid C# to provide a delegate for an object argument)
        /// </summary>
        [TestMethod]
        public void OutermostScopeFunctionMayNotBeTargetOfCallExpression()
        {
            var source = @"
				Set a = GetSomething.Name
				Function GetSomething()
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY101"]);
        }

        /// <summary>
        /// Very similar to OutermostScopeFunctionMayNotBeTargetOfCallExpression except that the function is within a class rather than in the outermost scope
        /// </summary>
        [TestMethod]
        public void ClassContainedFunctionMayNotBeTargetOfCallExpression()
        {
            var source = @"
				Class C1
					Function Go()
						Set a = GetSomething.Name
					End Function
					Function GetSomething()
					End Function
				End Class";
            TestCSharpCodeTranslationWithoutScaffolding(source, ["SKY103"]);
        }

        /// <summary>
        /// Some code was added to make common FOR loop structures more succinct in the generated C# (when an array is passed into a method ByRef and then the upper
        /// loop constraint is UBOUND of that array, for example) but that disabled ByRef argument mapping in cases where it shouldn't. This is a simple example where
        /// the ByRef x argument is passed to a builtin function, which will accept it ByVal and so there is no need for a ByRef mapping (which is required when an
        /// ByRef argument is passed to another method ByRef because that will require putting a reference to the original argument in a lambda, which is not legal
        /// in C#).
        /// </summary>
        [TestMethod]
        public void ByRefArgumentDoesNotRequireByRefArgumentMappingWhenPassedDirectlyToBuiltInFunction()
        {
            var source = @"
				Function F1(x)
					WScript.Echo TypeName(x)
				End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This is a companion to ByRefArgumentDoesNotRequireByRefArgumentMappingWhenPassedDirectlyToBuiltInFunction that illustrates that a ByRef mapping is
        /// required when a ByRef argument is passed to a builtin function if it is passed indirectly, via a nested function call.
        /// </summary>
        [TestMethod]
        public void ByRefArgumentRequireByRefArgumentMappingWhenPassedIndirectlyToBuiltInFunction()
        {
            var source = @"
				Function F1(x)
					WScript.Echo TypeName(F2(x))
				End Function

				Function F2(x)
					F2 = x
				End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This is another companion to ByRefArgumentDoesNotRequireByRefArgumentMappingWhenPassedDirectlyToBuiltInFunction - if a ByRef argument is passed to a builtin
        /// function and error-trapping may be enabled then a ByRef mapping will be required to avoid trying to reference the ref argument within the HANDLEERROR lambda
        /// </summary>
        [TestMethod]
        public void ByRefArgumentWillRequireByRefArgMapWhenPassedDirectlyToBuiltInFuncIfErrorMayBeOn() // ByRefArgumentWillRequireByRefArgumentMappingWhenPassedDirectlyToBuiltInFunctionIfErrorTrappingMayBeEnabled
        {
            var source = @"
				Function F1(x)
					On Error Resume Next
					WScript.Echo TypeName(x)
				End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This illustrated a bug that was identified with numeric literals of the form "&H001" - the trailing zeroes were causing an exception in the
        /// parsing process
        /// </summary>
        [TestMethod]
        public void EnsureThatHexValuesWithTrailingZeroesAreParsedCorrectly()
        {
            var source = @"
				const SOME_CONSTANT = &H0001
                Dim vv: vv = SOME_CONSTANT
			";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This proves a bug fix around the translation of statements within with blocks, where the first token in the statement is a member accessor - the CALL that
        /// was generated was incorrectly interpreting the method name as an argument
        /// </summary>
        [TestMethod]
        public void WithReferenceShouldNotConfuseBracketResolution()
        {
            var source = @"
				Function Render(x)
					With x
						.Draw ""Test""
					End With
				End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod, MyMemberData(nameof(VariousBracketDeterminedRefValArgumentData))]
        public void VariousBracketDeterminedRefValArgumentCases(int testNo, string source, string expectedResult)
        {
            TestCSharpCodeTranslationWithoutScaffolding(testNo, expectedResult, source, ["SKY101"]);
        }

        public static IEnumerable<object[]> VariousBracketDeterminedRefValArgumentData
        {
            get
            {
                yield return new object[] { 1, "func x", @"
            _.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Ref(_env.x, v => { _env.x = v; }));
" };
                yield return new object[] { 2, "func (x)", @"
            _.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Val(_env.x));
" };

                yield return new object[] { 3, "func x, y", @"
            _.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Ref(_env.x, v => { _env.x = v; }).Ref(_env.y, v2 => { _env.y = v2; }));
" };
                yield return new object[] { 4, "func (x), y", @"
            _.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Val(_env.x).Ref(_env.y, v => { _env.y = v; }));
" };
                yield return new object[] { 5, "func x, (y)", @"
            _.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Ref(_env.x, v => { _env.x = v; }).Val(_env.y));
" };
                yield return new object[] { 6, "z = func(x)", @"
            _env.z = _.VAL(_.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Ref(_env.x, v => { _env.x = v; })));
" };
                yield return new object[] { 7, "z = func(x, y)", @"
            _env.z = _.VAL(_.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Ref(_env.x, v => { _env.x = v; }).Ref(_env.y, v2 => { _env.y = v2; })));
" };
                yield return new object[] { 8, "z = func((x), y)", @"
            _env.z = _.VAL(_.CALLm0argp(this, _.NnO(_env.func, ""func""), _.ARGS.Val(_env.x).Ref(_env.y, v => { _env.y = v; })));
" };
            }
        }

        [TestMethod, MyMemberData(nameof(ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessaryData))]
        public void ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessary(int testNo, string source, string expectedResult)
        {
            TestCSharpCodeTranslationWithoutScaffolding(testNo, expectedResult, source, ["SKY101"]);
        }
        public static IEnumerable<object[]> ZeroArgumentBracketsEnforcedWhereAndOnlyWhereNecessaryData
        {
            get
            {
                yield return new object[] { 1, "a = b", @"_env.a = _.VAL(_env.b);" };
                yield return new object[] { 2, "a = b()", @"
            _env.a = _.VAL(_.CALLm0argp(this, _.NnO(_env.b, ""b""), _.ARGS.ForceBrackets()));
" };
                yield return new object[] { 3, "a = b(1)", @"
            _env.a = _.VAL(_.CALLm0argp(this, _.NnO(_env.b, ""b""), _.ARGS.Val((Int16)1)));
" };

                yield return new object[] { 4, "a = b.Name", @"
            _env.a = _.VAL(_.CALLm1v0(this, _.NnO(_env.b, ""b""), ""Name""));
" };
                yield return new object[] { 5, "a = b.Name()", @"
            _env.a = _.VAL(_.CALLm1argp(this, _.NnO(_env.b, ""b""), ""Name"", _.ARGS.ForceBrackets()));
" };
                yield return new object[] { 6, "a = b.Name(1)", @"
            _env.a = _.VAL(_.CALLm1v1(this, _.NnO(_env.b, ""b""), ""Name"", (Int16)1));
" };
            }
        }
    }
}
