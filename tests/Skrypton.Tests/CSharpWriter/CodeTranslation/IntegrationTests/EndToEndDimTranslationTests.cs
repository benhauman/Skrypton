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
            var expected = @"
                public object F1()
                {
                    object F1_retVal = null;
                    object myVariable = null;
                    return F1_retVal;
                }";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }

        [TestMethod, MyFact]
        public void DimWithDimensionsInsideFunction()
        {
            var source = @"
				Function F1()
					Dim myArray(63)
				End Function
			";
            var expected = @"
                public object F1()
                {
                    object F1_retVal = null;
                    object myArray = new object[64];
                    return F1_retVal;
                }";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
        }
    }
}
