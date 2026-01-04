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
            var TextBoxWebsite = new DialogGuiTextControl("TextBoxWebsite")
                //.InitializeTextControl("kuku")
                ;
            DoDialogGui(new Dictionary<string, object> { { TextBoxWebsite.ControlName, TextBoxWebsite } });
        }

        [TestMethod]
        public void CT35_LogChecklist_Dialog_388_OnSave() // 35:DFSnDLNeu
        {
            var externalReferences = new Dictionary<string, object> {
                { "TextBoxChecklist1URL", new DialogGuiTextControl("TextBoxChecklist1URL") },
                { "TextBoxChecklist2URL", new DialogGuiTextControl("TextBoxChecklist2URL") },
                { "TextBoxChecklist3URL", new DialogGuiTextControl("TextBoxChecklist3URL") },
                { "TextBoxChecklist4URL", new DialogGuiTextControl("TextBoxChecklist4URL") },
                { "TextBoxChecklist5URL", new DialogGuiTextControl("TextBoxChecklist5URL") },
                { "TextBoxChecklist6URL", new DialogGuiTextControl("TextBoxChecklist6URL") },
                { "TextBoxChecklist7URL", new DialogGuiTextControl("TextBoxChecklist7URL") },
                { "TextBoxChecklist8URL", new DialogGuiTextControl("TextBoxChecklist8URL") },
                { "TextBoxChecklist9URL", new DialogGuiTextControl("TextBoxChecklist9URL") },
                { "TextBoxChecklist10URL", new DialogGuiTextControl("TextBoxChecklist10URL") },
            };
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui, externalReferences);
            DoDialogGui(externalReferences);

        }

        private void DoDialogGui(Dictionary<string, object> externalReferences)
        {
            CncIn.ExecuteTranslatedProgram(TestCulture, TestContext.TestName, externalReferences);
        }
    }
}
