using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class DialogGui : TestBase
    {
        [TestMethod]
        public void QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click() => ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.DialogGui
            , new System.Collections.Generic.Dictionary<string, object>() { { "TextBoxWebsite", null } }
            );
    }
}
