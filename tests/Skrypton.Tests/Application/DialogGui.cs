using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.Tests.Application.Controls;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class DialogGui : TestBase
    {
        [TestMethod]
        public void QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click()// => TestDialogGui();
        //private void TestDialogGui()
        {
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui
                , new System.Collections.Generic.Dictionary<string, object>() { { "TextBoxWebsite", null } }
            );
            DoDialogGui();
        }

        private void DoDialogGui()
        {
            var TextBoxWebsite = new DialogGuiTextControl("TextBoxWebsite")
                //.InitializeTextControl("kuku")
                ;

            CncIn.ExecuteTranslatedProgram(TestCulture, TestContext.TestName, new Dictionary<string, object> { { "TextBoxWebsite", TextBoxWebsite } });
        }
    }
}
