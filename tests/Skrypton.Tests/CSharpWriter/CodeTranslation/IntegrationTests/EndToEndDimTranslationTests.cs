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
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }

        [TestMethod]
        public void DimWithDimensionsInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myArray(63)
				End Function
			";
            TestCSharpCodeTranslationWithoutScaffolding(null, source);
        }
    }
}
