using System;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public sealed class EndToEndRuntimeDateValidationTests : TestBase
    {
        /// <summary>
        /// If the only date literals can be safely validated at translation time and will not vary by culture, then there is no need to emit the ValidateAgainstCurrentCulture code
        /// </summary>
        [TestMethod, MyFact]
        public void NoRuntimeDateLiteralPresent()
        {
            var source = "If (a = #29 5 2015#) Then\nEnd If";
            TestCSharpCodeTranslation(source);
            //string expected = TextResourceHelper.LoadResourceText<TestBase>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension);
            //base.AreEqualStringArray(TestName, CSFileExtension,
            //    expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
            //    DefaultTranslator.Translate(TestCulture, source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable).Select(s => s.Content.Trim()).Where(s => s != "").ToArray()
            //);
        }

        /// <summary>
        /// If date literals are present in the source that need to be validated when the translated program is run (but before it does any other work), then extra code must be generated
        /// </summary>
        [TestMethod, MyFact]
        public void RuntimeDateLiteralPresent()
        {
            //TestCulture = CultureInfo.GetCultureInfo("en-GB");
            var source = "If (a = #29 May 2015#) Then\nEnd If";
            TestCSharpCodeTranslation(source);
            //
            //string expected = TextResourceHelper.LoadResourceText<TestBase>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension);
            //base.AreEqualStringArray(TestName, CSFileExtension,
            //    expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray(),
            //    DefaultTranslator.Translate(TestCulture, source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable).Select(s => s.Content.Trim()).Where(s => s != "").ToArray()
            //);
        }
    }
}
