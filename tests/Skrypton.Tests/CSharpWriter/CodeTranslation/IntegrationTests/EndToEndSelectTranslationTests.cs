using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndSelectTranslationTests : TestBase
    {
        /// <summary>
        /// This tests a fix made to select block translation - it was looking for token types based upon their content, rather than their type (so it was mistaking
        /// a StringToken whose content was a single comma characters as being an ArgumentSeparatorToken, if the type of the token is checked instead of its content
        /// then this sort of mistake will no longer occur)
        /// </summary>
        [TestMethod, MyFact]
        public void AllowSpecialCharactersToBeUsedAsStringsInSelectCases()
        {
            var source = @"
				Select Case x
					Case ""(""
						WScript.Echo ""Open""
					Case "")""
						WScript.Echo ""Close""
					Case "",""
						WScript.Echo ""Split""
				End Select";

            var expected = @"
				if (_.IF(_.EQ(_env.x, ""("")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Open"");
				}
				else if (_.IF(_.EQ(_env.x, "")"")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Close"");
				}
				else if (_.IF(_.EQ(_env.x, "","")))
				{
					_.CALLm1v1(this, _env.WScript, ""Echo"", ""Split"");
				}";
            TestCSharpCodeTranslationWithoutScaffolding(expected, source);
            //myAssert.AreEqual(
            //    expected.Replace(Environment.NewLine, "\n").Split(['\n'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray(),
            //    WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            //);
        }

		[TestMethod]
		public void SelectCaseWithStringTokens()
		{
            var source = @"
    Dim Size : Size = 0
	Suffix = "" B"" 
	Select Case Suffix 
		Case "" KB"" Size = Round(Size / 1024, 2) 
		Case "" MB""	Size = Round(Size / 1048576, 2) 
	End Select
";
			//_ = WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);

            base.TestCSharpCodeTranslation(source);
        }
    }
}
