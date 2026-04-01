using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndDoTranslationTests : TestBase
    {
        [TestMethod]
        public void SimpleDoWhile()
        {
            var source = @"
				DO WHILE i > 10
				LOOP
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void SimpleDoUntil()
        {
            var source = @"
				DO UNTIL i > 10
				LOOP
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void SimpleDoLoopWhile()
        {
            var source = @"
				DO
				LOOP WHILE i > 10
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void SimpleDoLoopUntil()
        {
            var source = @"
				DO
				LOOP UNTIL i > 10
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void DoLoopWithoutTerminationCondition()
        {
            var source = @"
				DO
				LOOP
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void ErrorHidingDoWhileLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO WHILE(1/0)
				LOOP
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void ErrorHidingDoUntilLoopWithErrorThrowingCondition()
        {
            var source = @"
				ON ERROR RESUME NEXT
				DO UNTIL(1/0)
				LOOP
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }
    }
}
