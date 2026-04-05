using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Helpline.Application.ScriptingModel;
using Microsoft.CodeAnalysis;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public sealed class ScriptControlTests : TestBase
    {
        [TestMethod] public void CT98__hlsysscript_cncIN() => DoCncInScriptControlTest();
        [TestMethod] public void DC_DATA__hlsysscript_cncIN() => DoCncInScriptControlTest();
        [TestMethod] public void LUNA12_quxDATA__hlsysscript_cncIN() => DoCncInScriptControlTest();
        private void DoCncInScriptControlTest()
        {
            CncInState state = new CncInState();
            DoScriptControlTest(state, state.CncInObjects, ["SKY102", "SKY106"], [], _ => { }, _ => { });
            Assert.IsNotNull(state.DoExtendWorkflowCaseIdentity, nameof(CncInState.DoExtendWorkflowCaseIdentity));
        }
        private sealed class CncInState
        {
            internal readonly Dictionary<string, object> CncInObjects;
            internal CncObj DoExtendWorkflowCaseIdentity = null;
            public CncInState()
            {
                Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext = Helpline.Application.ScriptingModel.ApplicationTestContext.Create(ctx =>
                {
                    ctx.HandlerMergeSUs = (obj) => { };
                });
                var session = Skrypton.Tests.Application.CncIn.CreateSampleConnectivityJob(cncTestContext);
                session.DoExtendWorkflowCaseOverride = (oi) =>
                {
                    DoExtendWorkflowCaseIdentity = (CncObj)oi;
                };

                CncInObjects = new Dictionary<string, object>() { { "session", session } };
            }
        }

        [TestMethod]
        public void Outlook1() => DoScriptControlTest<object>(null, [], [], [], services =>
        {
            services.RegisterHostService<IHostObjectFactoryHostService>(() => DialogGui.CreateTestHostObjectFactoryHostService()
                .RegisterObjectFactory<object>("Outlook.Application", (h) => DialogGui.CreateMyOutlookApplicationClass(h))
            );
            services.RegisterHostService<IHostMessageBoxHostService>(() => DialogGui.CreateHostMessageBoxHostService());
        }, (x) => { });

        [TestMethod]
        public void CT98_dialog276_OnLoad_ButtonCreate()
        {
            string source = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + ".vbs");
            Dictionary<string, ScriptExternalReferenceInfo> externalRefs = new Dictionary<string, ScriptExternalReferenceInfo> { { "ButtonCreate", new ScriptExternalReferenceInfo(new object(), []) } };
            TestCSharpCodeTranslation(source, ["ButtonCreate"], []);
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.Unknown, externalRefs: externalRefs);
            //CncIn.DoCncInTest(rsp);

            var externalReferences = externalRefs.ToDictionary(x => x.Key, v => v.Value.Instance);
            DoScriptControlTest<object>(null, externalReferences, [], [], services =>
            {
                //services.RegisterHostService<IHostObjectFactoryHostService>(() => DialogGui.CreateTestHostObjectFactoryHostService()
                //    .RegisterObjectFactory<object>("Outlook.Application", (h) => DialogGui.CreateMyOutlookApplicationClass(h))
                //);
                //services.RegisterHostService<IHostMessageBoxHostService>(() => DialogGui.CreateHostMessageBoxHostService());
            }, (x) => { });
        }

        // see 'ExecuteScriptByNameAsync'
        private void DoScriptControlTest<TState>(TState state, Dictionary<string, object> externalReferences, string[] translationSuppressions, string[] noWarn, Action<TestHostServices> setupHostServices, Action<TState> assertProgramOutput)
        {
            string chainName = TestName;

            string scriptContent = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".vbs");

            ScriptControlConfiguration controlConfig = CreateScriptControlConfiguration(false, translationSuppressions, noWarn);
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices(setupHostServices)), controlConfig);
            foreach (var externalReference in externalReferences)
            {
                scriptengine.AddObject(externalReference.Key, externalReference.Value); // https://jeffpar.github.io/kbarchive/kb/185/Q185697/
            }

            try
            {
                scriptengine.ExecuteStatement(scriptContent);
            }
            catch (CompilationFailedException ex)
            {
                Console.WriteLine(ex);
                throw;
            }

            assertProgramOutput(state);
        }
    }

    /*
TODO:
On Error Resume Next
    what, db, dialogid
ok:64bit! CreateObject("VBScript.RegExp") - not on linux! C:\Windows\SysWOW64\vbscript.dll   (on 64‑bit Windows); .tlb Embedded. TYPENAME(CreateObject("VBScript.RegExp")), IRegExp2
!!! CreateObject("Acceptance")       _CustomerTest_VRPayment21, 386
!!! CreateObject("ADODB.Command")    _CustomerTest_DeutschePost, 337
!!! CreateObject("ADODB.Connection") _CustomerTest_Mainova,	2
!!! CreateObject("ALA") _CustomerTest_Minfoline2021Mar,	297
!!! CreateObject("CardAccountVRP") _CustomerTest_VRPayment21	391
!!! CreateObject("ChangeTemplate") _CustomerTest_AYSTest 84
!!! CreateObject("DYMO.DymoAddIn") _CustomerTest_Webasto	3
!!! createobject("Excel.Application") _CustomerTest_Transcat	278
!!! CreateObject("helpLine.hlcontrols.HLHelperPFA	_CustomerTest_Gazprom	7
!!! CreateObject("Internetexplorer.application") 	_CustomerTest_PmcsHl2	276
!!! CreateObject("MAPI.Session") _CustomerTest_KVB	116
!!! CreateObject("Msxml2.DOMDocument") _CustomerTest_SwissGrid	349  => IXMLDOMDocument ("Msxml2.DOMDocument.3.0" or "Msxml2.DOMDocument.6.0", "Msxml2.FreeThreadedDOMDocument.3.0", "Msxml2.FreeThreadedDOMDocument.6.0"); CLSID_DOMDocument30, F6D90F11-9C73-11D3-B32E-00C04F990BB4/ 88d96a05-f192-11d4-a65f-0040963251e5	F6D90F12-9C73-11D3-B32E-00C04F990BB4/ 88d96a06-f192-11d4-a65f-0040963251e5, Header and IDL files (C/C++): msxml2.h, msxml2.idl, msxml6.h, msxml6.idl
!!! CreateObject("MSXML2.XMLHTTP") _CustomerTest_RatioData	692
!!! CreateObject("NetworkPort") objPort.SetValu	_CustomerTest_Storck	6
!!! CreateObject("Offsetting") _CustomerTest_Rhomberg	3
!!! CreateObject("Outlook.Application") _CustomerTest_PmcsHl2	1185
!!! CreateObject("Profile") _CustomerTest_SwissGrid	1694
!!! CreateObject("Scripting.FileSystemObject")  _CustomerTest_Gazprom	3
!!! CreateObject("Scriptlet.TypeLib") _CustomerTest_SwarovskiNeu	1425  => C:\Windows\System32\scriptlet.dll, ProgID: Scriptlet.TypeLib CLSID: {06290BD5-48AA-11D2-8432-006008C3FBFC} =>  Left(TypeLib.Guid, 38)
!!! CreateObject("System.Collections.ArrayList") _CustomerTest_BerlinerFw20211004AL	1316
!!! CreateObject("WbemScripting.SWbemDateTime")		_CustomerTest_DFSnDL	428
!!! CreateObject("WinHttp.WinHttpRequest.5.1")	_CustomerTest_HDM	448
!!! CreateObject("Word.Application")  _CustomerTest_BerlinerFw	393
!~! CreateObject("WScript.Shell") _CustomerTest_Tamedia	567
!!! CreateObject("Msxml2.ServerXMLHTTP.6.0")

!!! GetObject("winmgmts:") _CustomerTest_Tamedia	567 => WMI service connection => winmgmts is not a progid => use Set objLocator = CreateObject("WbemScripting.SWbemLocator") and then Set objWMI = objLocator.ConnectServer(".", "root\cimv2")
     */
}
