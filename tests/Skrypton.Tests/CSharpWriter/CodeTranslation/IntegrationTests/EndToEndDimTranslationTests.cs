using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndDimTranslationTests : TestBase
    {
        [TestMethod, MyFact]
        public void DimInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myVariable
				End Function
			";
            var expected = new[]
            {
                "public object f1()",
                "{",
                "object retVal1 = null;",
                "object myvariable = null;",
                "return retVal1;",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void DimWithDimensionsInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myArray(63)
				End Function
			";
            var expected = new[]
            {
                "public object f1()",
                "{",
                "object retVal1 = null;",
                "object myarray = new object[64];",
                "return retVal1;",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
