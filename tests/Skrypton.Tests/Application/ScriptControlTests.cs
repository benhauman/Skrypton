using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;
using System;
using System.Collections.Generic;
using System.Text;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class ScriptControlTests : TestBase
    {
        [TestMethod]
        public void DC_DATA__hlsysscript_cncIN()
        {
            DoScriptControlTest();
        }

        private void DoScriptControlTest() // see 'ExecuteScriptByNameAsync'
        {
            string chainName = TestName;

            //bool mergeSU_called = false;
            Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext = Helpline.Application.ScriptingModel.ApplicationTestContext.Create(ctx =>
            {
                ctx.HandlerMergeSUs = (obj) =>
                {
                    //mergeSU_called = true;
                };
            });
            var session = Skrypton.Tests.Application.CncIn.CreateSampleConnectivityJob(cncTestContext);

            //ExecuteTranslatedProgram(TestCulture, TestContext.TestName, new Dictionary<string, object> { { "session", session } });

            string scriptContent = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".vbs");

            IScriptControl scriptengine = new ScriptControlClass() { EngineCulture = TestCulture };
            scriptengine.Language = "VBScript";
            scriptengine.AllowUI = false;
            scriptengine.Timeout = -1;//MSScriptControl::NoTimeout;

            //scriptengine.AddObject(it.Key, objIDispatch, false); // https://jeffpar.github.io/kbarchive/kb/185/Q185697/
            scriptengine.AddObject("session", session);

            scriptengine.ExecuteStatement(scriptContent);

            /*
object[] args = new object[0];
string tmp = ScriptEnginePrefix + scriptName;
scriptControl.Run(tmp, args);//scriptControl.Run(ScriptEnginePrefix + scriptName, ref args);
             */
        }
    }
}
