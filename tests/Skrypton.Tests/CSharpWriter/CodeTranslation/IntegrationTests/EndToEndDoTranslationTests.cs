using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            TestCSharpCodeTranslationWithoutScaffoldingA(expected, source);
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
            TestCSharpCodeTranslationWithoutScaffoldingA(expected, source);
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
            TestCSharpCodeTranslationWithoutScaffoldingA(expected, source);
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
            TestCSharpCodeTranslationWithoutScaffoldingA(expected, source);
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
            TestCSharpCodeTranslationWithoutScaffoldingA(expected, source);
        }

        [TestMethod, MyFact]
        public void ErrorHidingDoWhileLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO WHILE(1/0)
				LOOP
			";
            var expected = @"
                int errOn = _.GETERRORTRAPPINGTOKEN();
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
                while (_.IF(() => _.IF(_.DIV((Int16)1, (Int16)0)), errOn))
                {
                }
                _.RELEASEERRORTRAPPINGTOKEN(errOn);";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        [TestMethod, MyFact]
        public void ErrorHidingDoUntilLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO UNTIL(1/0)
				LOOP
			";
            var expected = @"
                int errOn = _.GETERRORTRAPPINGTOKEN();
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
                while (_.IF(() => !_.IF(_.DIV((Int16)1, (Int16)0)), errOn))
                {
                }
                _.RELEASEERRORTRAPPINGTOKEN(errOn);";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }
    }
}
