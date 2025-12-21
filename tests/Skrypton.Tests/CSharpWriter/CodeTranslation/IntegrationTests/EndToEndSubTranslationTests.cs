using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
	public class EndToEndSubTranslationTests : TestBase
    {
        /// <summary>
        /// Since SUBs do not return values, attempting to set a return value within a SUB results in an illegal assignment error
        /// </summary>
        [TestMethod, MyFact]
        public void IfReturnValueSetPresentInSubThenRaiseIllegalAssignment()
        {
            var source = @"
                PUBLIC SUB F1()
                    F1 = Null
                END SUB
            ";
            var expected = new[]
            {
                "public void f1()",
                "{",
				"    _.SET(VBScriptConstants.Null, this, _.RAISEERROR(new IllegalAssignmentException(\"'F1'\")));",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If what looks like a return-value-setting statement is present within a SUB, but with brackets after the left-hand side NameToken, then it's
        /// a type mismatch error (which is the same with a FUNCTION.. but not a PROPERTY)
        /// </summary>
        [TestMethod, MyFact]
        public void IfReturnValueWithBracketsPresentInSubThenRaiseTypeMismatch()
        {
            var source = @"
                PUBLIC SUB F1()
                    F1() = Null
                END SUB
            ";
            var expected = new[]
            {
                "public void f1()",
                "{",
				"    _.SET(VBScriptConstants.Null, this, _.RAISEERROR(new TypeMismatchException(\"'F1'\")));",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture,source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
