using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public sealed class EndToEndRuntimeDateValidationTests : TestBase
    {
        /// <summary>
        /// If the only date literals can be safely validated at translation time and will not vary by culture, then there is no need to emit the ValidateAgainstCurrentCulture code
        /// </summary>
        [TestMethod]
        public void NoRuntimeDateLiteralPresent()
        {
            var source = "If (a = #29 5 2015#) Then\nEnd If";
            TestCSharpCodeTranslation(source);
        }

        /// <summary>
        /// If date literals are present in the source that need to be validated when the translated program is run (but before it does any other work), then extra code must be generated
        /// </summary>
        [TestMethod]
        public void RuntimeDateLiteralPresent()
        {
            var source = "If (a = #29 May 2015#) Then\nEnd If If (a = #02 June 2011#) Then\nEnd If";
            TestCSharpCodeTranslation(source);
        }
    }
}
