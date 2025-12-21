using Helpline.Application.ScriptingModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public class EblOnSave : TestBase
    {
        [TestMethod]
        public void db_TrumpfData_script_609_117971_5()
        {
            var oiDefault = new HLOBJECTID(494, 22222);
            ActionContext actx = new ActionContext() { LocaleId = 1026 };

            ActionArgs actargs = new ActionArgs() { m_oiDefault = oiDefault };
            EblContext hlContext = new EblContext(actx, actargs);
            EblObj objX = new EblObj(oiDefault);

            hlContext.LoadObject_Override = oi => objX;
            // CncIn.ExecuteTranslatedProgram(TestContext.TestName, new Dictionary<string, object> { { "hlcontext", hlContext } });
        }
    }
}
