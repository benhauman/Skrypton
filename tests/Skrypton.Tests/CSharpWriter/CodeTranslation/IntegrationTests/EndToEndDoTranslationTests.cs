using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndDoTranslationTests : TestBase
    {
        [TestMethod, MyFact]
        public void SimpleDoWhile()
        {
            var source = @"
				DO WHILE i > 10
				LOOP
			";
            var expected = new[]
            {
                "while (_.IF(_.GT(_.NullableNUM(_env.i), (Int16)10)))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void SimpleDoUntil()
        {
            var source = @"
				DO UNTIL i > 10
				LOOP
			";
            var expected = new[]
            {
                "while (!_.IF(_.GT(_.NullableNUM(_env.i), (Int16)10)))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void SimpleDoLoopWhile()
        {
            var source = @"
				DO
				LOOP WHILE i > 10
			";
            var expected = new[]
            {
                "do",
                "{",
                "} while (_.IF(_.GT(_.NullableNUM(_env.i), (Int16)10)));"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void SimpleDoLoopUntil()
        {
            var source = @"
				DO
				LOOP UNTIL i > 10
			";
            var expected = new[]
            {
                "do",
                "{",
                "} while (!_.IF(_.GT(_.NullableNUM(_env.i), (Int16)10)));"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void DoLoopWithoutTerminationCondition()
        {
            var source = @"
				DO
				LOOP
			";
            var expected = new[]
            {
                "while (true) { }"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void ErrorHidingDoWhileLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO WHILE(1/0)
				LOOP
			";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "while (_.IF(() => _.IF(_.DIV((Int16)1, (Int16)0)), errOn1))",
                "{",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void ErrorHidingDoUntilLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO UNTIL(1/0)
				LOOP
			";
            var expected = new[]
            {
                "var errOn1 = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn1);",
                "while (_.IF(() => !_.IF(_.DIV((Int16)1, (Int16)0)), errOn1))",
                "{",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn1);"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
