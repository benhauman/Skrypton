using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndErrorTrappingTests : TestBase
    {
        /// <summary>
        /// This is the most basic example - a single OnErrorResumeNext that applies to a single statement that follows it. Whenever any scope terminates,
        /// any error token must be released.
        /// </summary>
        [TestMethod, MyFact]
        public void SingleErrorTrappedStatement()
        {
            var source = @"
				On Error Resume Next
				WScript.Echo ""Test1""
			";
            var expected = @"
				var errOn = _.GETERRORTRAPPINGTOKEN();
				_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
				_.HANDLEERROR(errOn, () => {
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test1"");
				});
				_.RELEASEERRORTRAPPINGTOKEN(errOn);";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        /// <summary>
        /// If an error token is required, it will also be defined at the top of the scope, not just before the first OnErrorResumeNext (in case it is
        /// required elsewhere in the same VBScript scope but in a different C# block scope in the translated output)
        /// </summary>
        [TestMethod, MyFact]
        public void FlatStatementSetWithMiddleOneErrorTrapped()
        {
            var source = @"
				WScript.Echo ""Test1""
				On Error Resume Next
				WScript.Echo ""Test2""
				On Error Goto 0
				WScript.Echo ""Test3""
			";
            var expected = @"
				var errOn = _.GETERRORTRAPPINGTOKEN();
				_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test1"");
				_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
				_.HANDLEERROR(errOn, () => {
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test2"");
				});
				_.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
				_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test3"");
				_.RELEASEERRORTRAPPINGTOKEN(errOn);";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        /// <summary>
        /// Although the condition around the OnErrorResumeNext can never be met, the following statement will still have the error-trapping code around
        /// it since the analysis of the code paths only checks for what look like the potential to enable error-trapping. Since the condition is always
        /// false, in the translated code the STARTERRORTRAPPING call will not be made and so the HANDLEERROR will perform no work (it will not trap the
        /// error) but it is a layer of redirection that can be avoided if the translator is sure that it's not required (if there was no OnErrorResumeNext
        /// present all, for example).
        /// </summary>
        [TestMethod, MyFact]
        public void ErrorTrappingLayerMustBeAddedEvenIfItWillOnlyPotentiallyBeEnabled()
        {
            var source = @"
				If (False) Then
					On Error Resume Next
				End If
				WScript.Echo ""Test1""
			";
            var expected = @"
				var errOn = _.GETERRORTRAPPINGTOKEN();
				if (_.IF(false))
				{
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
				}
				_.HANDLEERROR(errOn, () => {
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test1"");
				});
				_.RELEASEERRORTRAPPINGTOKEN(errOn);";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        [TestMethod, MyFact]
        public void ErrorTrappingDoesNotAffectChildScopes()
        {
            var source = @"
				On Error Resume Next
				Func1
				Function Func1()
					WScript.Echo ""Test1""
				End Function
			";
            var expected = @"
				var errOn = _.GETERRORTRAPPINGTOKEN();
				_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
				_.HANDLEERROR(errOn, () => {
					_.CALLm1v0(this, _outer, ""Func1"");
				});
				_.RELEASEERRORTRAPPINGTOKEN(errOn);
				public object Func1()
				{
					object Func1_retVal = null;
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test1"");
					return Func1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        [TestMethod, MyFact]
        public void ErrorTrappingDoesNotAffectParentScopes()
        {
            var source = @"
				Func1
				WScript.Echo ""Test2""
				Function Func1()
					On Error Resume Next
					WScript.Echo ""Test1""
				End Function
			";
            var expected = @"
				_.CALLm1v0(this, _outer, ""Func1"");
				_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test2"");
				public object Func1()
				{
					object Func1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					_.HANDLEERROR(errOn, () => {
						_.CALLm1v1(this, _env.WScript, ""Echo"", ""Test1"");
					});
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return Func1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        /// <summary>
        /// The "Raise" and "Clear" methods on the VBScript "Err" reference need to be remapped onto the support class functions RAISEERROR and CLEARANYERROR.
        /// This may be done directly if supported numbers of arguments are present - if not then the support function will have to be called through the CALL
        /// method so that the invalid argument count results in a runtime error rather than a compile failure (the C# will not compile if the functions are
        /// specified as being called directly but with incorrect argument counts), in order to be consistent with VBScript.
        /// </summary>
        [TestMethod, MyFact]
        public void TranslateErrRaiseIntoAppropriateSupportFunction()
        {
            var source = @"
				Err.Raise vbObjectError
				Err.Raise vbObjectError, ""Source""
				Err.Raise vbObjectError, ""Source"", ""Test""
				Err.Raise vbObjectError, ""Source"", ""Test"", ""Bonus Argument""
				Err.Clear
				Err.Clear()
				Err.Clear ""Bonus Argument""
			";
            var expected = @"
				_.RAISEERROR(VBScriptConstants.vbObjectError);
				_.RAISEERROR(VBScriptConstants.vbObjectError, ""Source"");
				_.RAISEERROR(VBScriptConstants.vbObjectError, ""Source"", ""Test"");
				_.CALLm1argp(this, _, ""RAISEERROR"", _.ARGS.Val(VBScriptConstants.vbObjectError).Val(""Source"").Val(""Test"").Val(""Bonus Argument""));
				_.CLEARANYERROR();
				_.CLEARANYERROR();
				_.CALLm1v1(this, _, ""CLEARANYERROR"", ""Bonus Argument"");";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        /// <summary>
        /// If a by-ref argument of the containing function is accessed in code that may be wrapped in a lambda (ie. within a HANDLEERROR call) then an alias
        /// for the by-ref argument must be used, since it's not valid C# to access a ref argument within a lambda. If there is a possibility that this by-ref
        /// argument's value will be changed within the lambda then the alias value must be copied back over the by-ref argument value after the lambda's work
        /// has completed (even if it errors, in case the value was updated before the error occurred). However, there are places where it's not possible for
        /// the by-ref argument to be changed within the lambda - if "a" is a by-ref argument and the work inside the lambda is "F1(a + 1)" then there's no
        /// way that F1 could affect "a", and so the copy-back-from-alias code is not emitted. There was a flaw in the logic that meant that if a by-ref
        /// argument was the target of a value-setting statement (eg. "a = F1", where "a" is a by-ref argument of the containing function) then it would not
        /// be identified as needing to be overwritten by the alias value after the value-setting statement is processed within the lambda. This meant that
        /// the right-hand-side of the value-setting statement would be evaluated, but its return value would not be applied to the by-ref argument.
        /// </summary>
        [TestMethod, MyFact]
        public void IfValueSettingStatementTargetIsByRefArgumentOfContainingFunctionThenAnyByRefMappingMustBeTwoWay()
        {
            // 2016-03-09 DWR: It's important that the ON ERROR RESUME NEXT comes after the value setting statement since the translation process is not
            // currently clever enough to determine that this means that error-trapping could never apply to the value setting statement - the fact that
            // there IS some error-handling within the function scope means that the translation process goes on the assumption that it's possible that
            // the error-handling might impact the value-setting (this is something that could be improved in the future since it would result in more
            // succinct code in places where there is error-handling within the same scope but which it can be proved can never apply to a particular
            // statement).
            var source = @"
				Function F1(x)
					Set x = Nothing
					On Error Resume Next
				End Function
			";
            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					object byrefalias = x;
					try
					{
						byrefalias = VBScriptConstants.Nothing;
					}
					finally { x = byrefalias; }
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// When the by-ref argument aliasing logic was fixed such that the test
        ///   IfValueSettingStatementTargetIsByRefArgumentOfContainingFunctionThenAnyByRefMappingMustBeTwoWay
        /// could be passed, value setting statement targets were getting aliased when they didn't need to. The particular case that was being addressed
        /// was when the value setting statement existed within a scope that may require error-trapping. This test ensures that the over-aggressive
        /// aliasing is no longer applied.
        /// </summary>
        [TestMethod, MyFact]
        public void ValueSettingTargetThatIsByRefFunctionArgumentShouldOnlyBeReadWriteAliasedIfWithinErrorHandling()
        {
            var source = @"
				Function F1(x)
					Set x = Nothing
				End Function
			";
            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					x = VBScriptConstants.Nothing;
					return F1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// An IF block within a FUNCTION that has an OERN before the block must have error-trapping around its inner statements
        /// </summary>
        [TestMethod, MyFact]
        public void IfBlockStatementsMustBeWrappedInErrorHandlingIfWithinFunctionWithOnErrorResumeNextBeforeIfBlock()
        {
            var source = @"
				Function F1(ByVal value)
					On Error Resume Next
					If True Then
						F1 = DateValue(value)
					End If
					On Error Goto 0
				End Function
			";
            var expected = @"
				public object F1(object value)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					if (_.IF(() => true, errOn))
					{
						_.HANDLEERROR(errOn, () => {
							F1_retVal = _.DATEVALUE(value);
						});
					}
					_.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// This is related to IfBlockStatementsMustBeWrappedInErrorHandlingIfWithinFunctionWithOnErrorResumeNextBeforeIfBlock - if the OERN comes after
        /// the IF block then it will not affect the IF block and so the IF block does not need to consider any error-trapping (around either its own
        /// conditional codeExpression or its inner statements)
        /// </summary>
        [TestMethod, MyFact]
        public void IfBlockStatementsNeedNotBeWrappedInErrorHandlingIfWithinFunctionWithOnErrorResumeNextAfterIfBlock()
        {
            var source = @"
				Function F1(ByVal value)
					If True Then
						F1 = DateValue(value)
						Exit Function
					End If
					On Error Resume Next
					F1 = Date()
				End Function
			";
            var expected = @"
				public object F1(object value)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					if (_.IF(true))
					{
						F1_retVal = _.DATEVALUE(value);
						_.RELEASEERRORTRAPPINGTOKEN(errOn);
						return F1_retVal;
					}
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					_.HANDLEERROR(errOn, () => {
						F1_retVal = _.DATE();
					});
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        /// <summary>
        /// This is related to IfBlockStatementsNeedNotBeWrappedInErrorHandlingIfWithinFunctionWithOnErrorResumeNextAfterIfBlock - if the OERN comes after
        /// the IF block but both it and the IF block are within a looping structure (such as a FOR block) then the IF block *does* need error handling
        /// because there is a chance that the second pass through the loop could occur with error-trapping enabled.
        /// conditional codeExpression or its inner statements)
        /// </summary>
        [TestMethod, MyFact]
        public void IfBlockStatementsNeedsToBeWrappedInErrorHandlingIfOnErrorResumeNextAfterComesAfterItWithinLoopingStructure()
        {
            // The code analysis is not clever enough to realise that the FOR block will only be executed once (since it starts and ends at 1) and so it
            // presumes that it will be executed multiple times and so the IF block needs to be able to handle error-trapping
            var source = @"
				Function F1(ByVal value)
					Dim i: For i = 1 To 1
						If (True) Then
							F1 = DateValue(value)
						End If
						On Error Resume Next
					Next
				End Function
			";
            var expected = @"
				public object F1(object value)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					object i = null;
					i = (Int16)1;
					while (true)
					{
						if (_.IF(() => true, errOn))
						{
							_.HANDLEERROR(errOn, () => {
								F1_retVal = _.DATEVALUE(value);
							});
						}
						_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
						var continueLoop = false;
						_.HANDLEERROR(errOn, () => {
							i = _.ADD(i, (Int16)1);
							continueLoop = _.StrictLTE(i, 1);
						});
						if (!continueLoop)
							break;
					}
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

        private static IEnumerable<string> SplitOnNewLinesSkipFirstLineAndTrimAll(string value)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            return value.NormalizeLineEndings().SplitLines().Skip(1).Select(v => v.Trim());
        }
    }
}
