using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Helpline.Application.ScriptingModel;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public class CncIn : TestBase
    {
        /*
            SELECT scriptid   = scr.id
                , scripttext = scr.script
             FROM [dbo].[hlsysscript] AS scr
            WHERE scr.active      = 1
              AND scr.objectdefid = 900 -- 900:connectivity
              AND scr.[type]      = 16 -- 16:ScriptTypeConnectivityIn
              AND scr.scriptmode  = 0 -- 0:eScriptMode.ScriptModeWorking
              AND LEN(ISNULL(scr.script,N'')) > 0 -- TODO: CK_
              ;
         */
        /* --> row number is the customer index : 98:PsoShow, 35:DFSnDLNeu
        SELECT [dbname]
               ,[sizebytes_before]
               ,[sizebytes_after]
               ,[hasfilestreamfilegroup]
               ,[sizegb_before]
               ,[sizegb_after]
               ,[deltamb]
           FROM [CustomerAnalytics].[dbo].[_DatabaseStats]

         */
        [TestMethod]
        public void DC_DATA__hlsysscript_cncIN()
        {
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.Connectivity, externalRefs: ChainsTest.CollectExternalRefs(ScriptUsageKind.Connectivity), suppressions: ["SKY102", "SKY104", "SKY106"]);
            DoCncInTest(rsp);
        }
        [TestMethod]
        public void LUNA12_quxDATA__hlsysscript_cncIN()
        {
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.Connectivity, externalRefs: ChainsTest.CollectExternalRefs(ScriptUsageKind.Connectivity), suppressions: ["SKY102", "SKY104", "SKY106"]);
            DoCncInTest(rsp);
        }
        [TestMethod]
        public void CT98__hlsysscript_cncIN()
        {
            TestScriptResponse rsp = ChainsTest.TestScriptChain(this, ScriptUsageKind.Connectivity, externalRefs: ChainsTest.CollectExternalRefs(ScriptUsageKind.Connectivity));
            DoCncInTest(rsp);
        }

        private void DoCncInTest(TestScriptResponse rsp)
        {
            bool mergeSU_called = false;
            Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext = Helpline.Application.ScriptingModel.ApplicationTestContext.Create(ctx =>
            {
                ctx.HandlerMergeSUs = (obj) =>
                {
                    mergeSU_called = true;
                };
            });
            CncJob session = CreateSampleConnectivityJob(cncTestContext);
            CncObj DoExtendWorkflowCaseIdentity = null;
            session.DoExtendWorkflowCaseOverride = (oi) =>
            {
                DoExtendWorkflowCaseIdentity = (CncObj)oi;
            };
            var hostServices = CreateTestHostServices();
            //string translated_cs_expected = rsp.TranslatedCsCode;// TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension);
            ExecuteTranslatedProgram(this, rsp.TranslatedCsCode, hostServices, new Dictionary<string, ScriptExternalReferenceInfo> { { "session", new ScriptExternalReferenceInfo(session, []) } }, gr => { });

            // assert
            Assert.IsFalse(mergeSU_called, "mergeSU_called");
            Assert.IsNotNull(DoExtendWorkflowCaseIdentity, nameof(DoExtendWorkflowCaseIdentity));

        }

        internal static void ExecuteTranslatedProgram(TestBaseX tst, string translatedCsCode, IServiceProvider hostServices, IReadOnlyDictionary<string, ScriptExternalReferenceInfo> externalReferences, Action<GlobalReferencesBase> dialogHandler)
        {
            IRuntimeHost runtimeHost = new TestRuntimeHost(hostServices);
            var scriptControlClass = tst.CreateScriptControlClass(runtimeHost, tst.CreateScriptControlConfiguration(false, [], []));

            scriptControlClass.TestSetDefaultRuntimeFunctionalityProviderSetup((x) => SetupDefaultRuntimeFunctionalityProvider(x, hostServices, tst.TestCulture));

            //RunTranslatedProgram(scriptengineClass, runtimeLogger, hostServices, externalReferences, dialogHandler);
            foreach (KeyValuePair<string, ScriptExternalReferenceInfo> externalReferencesEntry in externalReferences)
            {
                string externalReferenceName = externalReferencesEntry.Key;
                ScriptExternalReferenceInfo nfo = externalReferencesEntry.Value;
                object externalReferenceInstance = nfo.Instance;
                IScriptControl scriptControl = scriptControlClass;
                scriptControl.AddObject(externalReferenceName, externalReferenceInstance, nfo.AddMembers);
            }

            scriptControlClass.TestTranslatedStatement(tst.TestName, translatedCsCode, ["CS0219"], doRun: true, dialogHandler);
        }
        internal static void RunTranslatedProgram(IRuntimeLogger runtimeLogger, IServiceProvider hostServices, CultureInfo culture, Type tRunner, IReadOnlyDictionary<string, object> externalReferences, Action<GlobalReferencesBase> dialogHandler)
        {
            IRuntimeHost runtimeHost = new TestRuntimeHost(hostServices);
            DefaultRuntimeSupportClassFactory defaultRuntimeSupportClassFactoryInstance = Skrypton.RuntimeSupport.DefaultRuntimeSupportClassFactory.Create(runtimeHost, runtimeLogger, culture);
            DefaultRuntimeFunctionalityProvider compatLayer = new DefaultRuntimeFunctionalityProvider(runtimeHost, runtimeLogger, defaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever, culture);
            SetupDefaultRuntimeFunctionalityProvider(compatLayer, hostServices, culture);

            RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(tRunner, compatLayer);

            EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();

            var properties = environmentReferences.GetType().GetProperties();
            var propertiesNameInfo = properties.ToDictionary(x => x.Name, x => x, StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, object> externalReferencesEntry in externalReferences)
            {
                string externalReferenceName = externalReferencesEntry.Key;
                object externalReferenceInstance = externalReferencesEntry.Value;
                environmentReferences.InitializeExternalReference(externalReferenceName, externalReferenceInstance);

                if (!propertiesNameInfo.TryGetValue(externalReferenceName, out PropertyInfo pi_externalReference1))
                    throw new InvalidOperationException($"Invalid external reference '{externalReferenceName}'.");
                // sanity check
                _ = pi_externalReference1.GetValue(environmentReferences);
            }

            GlobalReferencesBase gr = runner.Run(environmentReferences);
            dialogHandler(gr);
        }

        internal static void SetupDefaultRuntimeFunctionalityProvider(DefaultRuntimeFunctionalityProvider provider, IServiceProvider hostServices, CultureInfo culture)
        {
            provider.RegisterObjectCreateFactory("Scripting.Dictionary", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyScriptingDictionaryCpuAny());
            provider.RegisterObjectCreateFactory("Shell.Application", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyShellApplication());
            provider.RegisterObjectCreateFactory("Msxml2.ServerXMLHTTP.6.0", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyServerXMLHTTP60());
            provider.RegisterObjectCreateFactory("Msxml2.DOMDocument", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyMsxml2DOMDocument());
            provider.RegisterObjectCreateFactory("VBScript.RegExp", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyVBScriptRegExp(culture));
            provider.RegisterObjectCreateFactory("WScript.Shell", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyWScriptShell(hostServices));
            provider.RegisterObjectCreateFactory("WbemScripting.SWbemLocator", (optionalMonikerValues) => new Skrypton.Tests.RuntimeSupport.Implementations.MySWbemLocator(hostServices, optionalMonikerValues));
            //provider.RegisterObjectCreateFactory("ADODB.Connection", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.ADODB.MyADODBConnection(hostServices));
            //provider.RegisterObjectCreateFactory("ADODB.Connection", (_) =>  DialogGui.CreateADODBConnectionClass( Skrypton.Tests.RuntimeSupport.Implementations.ADODB.MyADODBConnection(hostServices));
            provider.RegisterObjectCreateFactory("ADODB.Command", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.ADODB.MyADODBCommand());
            provider.RegisterObjectCreateFactory("ADODB.Recordset", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.ADODB.MyADODBRecordSet());
            provider.RegisterObjectCreateFactory("ADODB.Stream", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.ADODB.MyADODBStream());
            provider.RegisterObjectCreateFactory("Scripting.FileSystemObject", (_) => Skrypton.Tests.RuntimeSupport.Components.FileSystemSupport.MyFileSystemObject.Create(hostServices));
            provider.RegisterObjectCreateFactory("Scriptlet.TypeLib", (_) => new Skrypton.Tests.RuntimeSupport.Implementations.MyScriptletTypeLib());
        }


        internal static Helpline.Application.ScriptingModel.CncJob CreateSampleConnectivityJob(Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext)
        {
            return new CncJob(cncTestContext)
            {
                m_cfg = new CncConfigGroup("Root").AddGroup("casetypEs", caseTypes =>
                {
                    caseTypes.AddGroup("type1", t1 =>
                    {
                        t1.InitValue("CaseType", v => { v.m_data = null; });
                        t1.InitValue("MailAttributeKey", v => { v.m_data = "PersonCommunication.PersonEmail_CA.EmailAddress"; });
                        t1.InitValue("Type", v => { v.m_data = "1"; });
                    });
                    caseTypes.AddGroup("type2", t2 =>
                    {
                        t2.InitValue("CaseTyp", v => { v.m_data = null; });
                        t2.InitValue("MailAttributeKey", v => { v.m_data = "PersonCommunication.PersonEmail_CA.EmailAddress"; });
                        t2.InitValue("Type", v => { v.m_data = "-2"; });
                    });
                })
                                    ,

                m_mailRequest = new CncMail()
                {
                    Subject = "this a feedbacl [#20190711-0012]. Awesome",
                    data_From = "peter.pan@wonderland.com"
                }
            };
        }
    }
}
