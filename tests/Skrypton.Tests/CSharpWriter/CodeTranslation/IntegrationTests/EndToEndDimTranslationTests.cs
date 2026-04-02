using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndDimTranslationTests : TestBase
    {
        [TestMethod]
        public void DimInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myVariable
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void DimWithDimensionsInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myArray(63)
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        //[Ignore]
        public void DimSpaceX() // CT98_GlobalScript
        {
            var source = @"Dim i, CharCode, Char, Space, URLEncode
				Space = ""+""
                URLEncode = Space
                URLEncode = Space & ""x"" & Space
                URLEncode = ""y"" & Space
                URLEncode = Space & ""z""
                URLEncode = F1(Space)
                URLEncode = F2(Space)
                URLEncode = F3(Space)
                URLEncode = F4(Space)
                URLEncode = F5(Space)
                Function F1()
					Dim spaCe
				End Function
                Function F2(SPace)
					F2 = space
				End Function
                Function F3(ByRef SPace)
					F3 = space
				End Function
                Function F4(ByVal SPACE)
					F4 = space
				End Function
                Function F5(ByRef txt)
					F5 = txt & SPACE(77)
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, []);

            //Assert.Inconclusive("Argument not valid");
            // actual: _outer.URLEncode = _.CONCAT(_.CALLm1v0(this, _, "SPACE"), "x", _.CALLm1v0(this, _, "SPACE"));
            // expect: _outer.URLEncode = _.CONCAT( ..._outer.Space
        }
        [TestMethod]
        public void DimSpace1() // see 'TranslateCallExpressionSegment'
        {
            var source = @"Dim CharCode, Space, URLEncode
                URLEncode = Space
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, []);
        }
        [TestMethod]
        public void DimSpace2() // see 'TranslateCallExpressionSegment'
        {
            var source = @"Dim CharCode, Space, URLEncode
                URLEncode = Space & CharCode
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, []);
        }
        [TestMethod]
        public void DimSpace3() // see 'TranslateCallExpressionSegment'
        {
            var source = @"Dim CharCode, Space, URLEncode
                Function F5(ByRef txt)
					F5 = txt & SPACE(77)
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, []);
        }
        [TestMethod]
        public void DimSpace4() // see 'TranslateCallExpressionSegment'
        {
            var source = @"Dim txt
txt = Space(10)
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, [], [], []);
        }
        [TestMethod]
        public void DimSpace5() // see 'TranslateCallExpressionSegment'
        {
            var source = @"Dim URLEncode
                URLEncode = Space(10)
                Function F5(ByRef txt)
                    Dim Space: Space = ""__""
					F5 = txt & Space
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(source, [], [], []);
        }
    }
}
