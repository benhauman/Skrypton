using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Helpline.Application.ScriptingModel;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;
using Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Skrypton.RuntimeSupport;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Application
{
    // D:\zapechene.2015\VBScript.Parse\LuboVBParser1\TestResources
    [TestClass]
    public sealed class ChainsTest : TestBase
    {
        [TestMethod, MyMemberData(nameof(ChainNames))]
        public void Chains(string chainName, ScriptUsageKind scriptUsage)
        {
            if (chainName == "CT125_ClientComputer_Dialog_349_ButtonGeneralInfo_Click"
             || chainName == "CT130_ClientComputer_Dialog_567_Button1_Click"
             || chainName == "CT74_ClientComputer_Dialog_2_ButtonShowWebsite_Click"
             || (chainName.StartsWith("CT") && chainName.Contains("_Dialog", StringComparison.Ordinal))
             )
            {
                // ignore for now: the undeclared external references  should be rendered as environment references and not a variables in 'Go'
                return;
            }
            MemberDataTestName = chainName;

            TestScriptResponse rsp = TestScriptChain(this, scriptUsage);
            var tst = this;
            var hostServices = CreateTestHostServices();
            var externalReferences = new Dictionary<string, object>();

            if (scriptUsage == ScriptUsageKind.EBL)//(scriptContent.Contains("hlContext"))
            {
                var oiDefault = new HLOBJECTID(494, 22222);
                ActionContext actx = new ActionContext() { LocaleId = 1026 };

                ActionArgs actargs = new ActionArgs() { m_oiDefault = oiDefault };
                EblContext hlContext = new EblContext(actx, actargs);
                EblObj objX = new EblObj(oiDefault);
                hlContext.LoadObject_Override = oi => objX;

                externalReferences.Add("hlContext", hlContext); // EBL
            }
            else if (scriptUsage == ScriptUsageKind.Connectivity)
            {
                Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext = Helpline.Application.ScriptingModel.ApplicationTestContext.Create(ctx =>
                {
                });
                CncJob session = CncIn.CreateSampleConnectivityJob(cncTestContext);
                session.DoExtendWorkflowCaseOverride = (oi) => { };
                externalReferences.Add("session", session); // Connectivity IN/OUT
            }
            else if (scriptUsage == ScriptUsageKind.DialogGui)
            {
            }

            var scriptengineClass = CreateScriptControlClass(new TestRuntimeHost(hostServices));
            scriptengineClass.TestTranslatedStatement(rsp.TranslatedCsCode, [
                "CS0219", // error CS0219: The variable 'ForWriting' is assigned but its value is never used
                ], doRun: false, gr => { });
        }

        public static object[][] ChainNames
        {
            get
            {
                List<string> names = new List<string>();

                Assembly resourceAssembly = typeof(CncIn).Assembly;
                string[] resource_names = resourceAssembly.GetManifestResourceNames()
                    .OrderBy(x => x).ToArray();

                string prefix = "Skrypton.Tests.VbsResources.";
                string suffix = ".vbs";
                foreach (string resAsm_name in resource_names)
                {
                    if (resAsm_name.StartsWith(prefix))
                    {
                        if (resAsm_name.EndsWith(suffix))
                        {
                            if (resAsm_name.EndsWith(".generated.vbs"))
                            {

                            }
                            else if (resAsm_name.EndsWith("_DialogGlobalScript.vbs"))
                            {

                            }
                            else
                            {
                                names.Add(resAsm_name.Substring(prefix.Length, resAsm_name.Length - prefix.Length - suffix.Length)); // ".vbs"
                            }
                        }
                    }
                }

                List<object[]> result = new List<object[]>();
                foreach (string chainName in names)
                {
                    bool isCnc = chainName.Contains("_cncIN", StringComparison.OrdinalIgnoreCase) || chainName.Contains("_900_");
                    bool isDialog = chainName.Contains("_Dialog", StringComparison.OrdinalIgnoreCase) || chainName.Contains("_Web", StringComparison.OrdinalIgnoreCase);
                    //bool isEBL = chainName.Contains("_EBL", StringComparison.OrdinalIgnoreCase);
                    ScriptUsageKind scriptUsage = isCnc
                        ? ScriptUsageKind.Connectivity
                        : isDialog
                            ? ScriptUsageKind.DialogGui
                            : ScriptUsageKind.EBL;
                    //scriptContent.Contains("hlContext")

                    result.Add([chainName, scriptUsage]);
                }

                return result.ToArray();
            }
        }
        public static TestScriptResponse TestScriptChain(TestBaseX tst, ScriptUsageKind scrUsage, IReadOnlyDictionary<string, object> externalRefs = null, bool isOptionalAssert = false)
        {
            string chainName = tst.TestName;
            string scriptContent = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".vbs");
            string generated_vbs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".generated.vbs", isOptionalAssert);
            string translated_cs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + CSFileExtension, isOptionalAssert);
            string xml_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".xml", isOptionalAssert);

            string customerDialogGlobalScript;
            if (scrUsage == ScriptUsageKind.DialogGui || scrUsage == ScriptUsageKind.DialogWeb)
            {
                string[] chainTokens = chainName.Split('_');
                string customerAlias = chainTokens[0];
                if (!customerAlias.StartsWith("CT", StringComparison.OrdinalIgnoreCase)) // 'QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click' => QUX_HLData
                {
                    customerAlias = $"{chainTokens[0]}_{chainTokens[1]}";
                }
                customerDialogGlobalScript = TextResourceHelper.LoadResourceText<CncIn>($"Skrypton.Tests.VbsResources.{customerAlias}_DialogGlobalScript.vbs"); // see [hlsysdialogglobalscript]
            }
            else
            {
                customerDialogGlobalScript = null;
            }

            return TestScriptChainX(tst, chainName,
                customerDialogGlobalScript,
                scriptContent,
                generated_vbs_expected,
                translated_cs_expected,
                xml_expected,
                scrUsage,
                externalRefs, isOptionalAssert);
        }
        public static TestScriptResponse TestScriptChainX(TestBaseX tst, string chainName,
                string customerDialogGlobalScript,
                string scriptContent,
                string generated_vbs_expected,
                string translated_cs_expected,
                string xml_expected,
                ScriptUsageKind scrUsage,
                IReadOnlyDictionary<string, object> externalRefs = null,
                bool isOptionalAssert = false)
        {
            NonNullImmutableList<string> externalDependencies = new NonNullImmutableList<string>();
            if (externalRefs == null)
            {
                if (scrUsage == ScriptUsageKind.EBL)//(scriptContent.Contains("hlContext"))
                    externalDependencies = externalDependencies.Add("hlContext"); // EBL
                if (scrUsage == ScriptUsageKind.Connectivity)
                    externalDependencies = externalDependencies.Add("session"); // Connectivity IN/OUT
            }
            else
            {
                foreach (string externalRefName in externalRefs.Keys)
                {
                    externalDependencies = externalDependencies.Add(externalRefName);
                }
            }

            if (scrUsage == ScriptUsageKind.DialogGui || scrUsage == ScriptUsageKind.DialogWeb)
            {
                string[] chainTokens = chainName.Split('_');
                string customerAlias = chainTokens[0];
                if (!customerAlias.StartsWith("CT", StringComparison.OrdinalIgnoreCase)) // 'QUX_HLData_Contact_Dialog_2_ButtonShowWebsite_Click' => QUX_HLData
                {
                    customerAlias = $"{chainTokens[0]}_{chainTokens[1]}";
                }

                StringBuilder completeDialogScript = new StringBuilder();
                if (!string.IsNullOrEmpty(customerDialogGlobalScript))
                {
                    if (completeDialogScript.Length > 0)
                    {
                        completeDialogScript.AppendLine();
                    }
                    completeDialogScript.Append(customerDialogGlobalScript);
                }

                if (!string.IsNullOrEmpty(scriptContent))
                {
                    if (completeDialogScript.Length > 0)
                    {
                        completeDialogScript.AppendLine();
                    }
                    completeDialogScript.Append(scriptContent);
                }

                scriptContent = completeDialogScript.ToString();
            }

            //Console.WriteLine("parsing...");
            var parsed_items = Skrypton.LegacyParser.Parser.Parse(tst.TestCulture, scriptContent);

            StringBuilder parsed_output = new StringBuilder();
            var generationContext = BaseSourceGenerationContextDefault.CreateBaseSourceGenerationContext();
            foreach (ICodeBlock parsedBlock in parsed_items)
            {
                parsed_output.AppendLine(parsedBlock.GenerateBaseSource(generationContext));
            }

            string workItemName = "Script";// TestContext.TestName;
            string generated_vbs_actual = parsed_output.ToString().NormalizeLineEndings();

            string failed_text = null;
            string storedFile;

            if (generated_vbs_expected != null)
            {
                if (generated_vbs_expected != generated_vbs_actual)
                {
                    storedFile = tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".generated.vbs", generated_vbs_expected, generated_vbs_actual);
                    failed_text = "VBS generation failed. See 'Output' for more information. storedFile:" + storedFile;
                }
            }

            var outermostBlock = Skrypton.LegacyParser.Parser.ParseToOutermostScope(parsed_items);
            string xml_actual = ToXml(outermostBlock, x => failed_text = x);

            if (xml_expected != null)
            {
                if (xml_expected != xml_actual)
                {
                    storedFile = tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".xml", xml_expected, xml_actual);
                    failed_text = "Xml generation failed. See 'Output' for more information. storedFile:" + storedFile;
                }
            }


            Console.WriteLine("translating...");
            string translated_cs_actual = DefaultCSharpTranslation.GetTranslatedProgramCode(tst.TestCulture, scriptContent, externalDependencies);

            //IEnumerable<TranslatedStatement> translated_items = Skrypton.CSharpWriter.DefaultTranslator.Translate(tst.TestCulture, scriptContent, externalDependencies.ToArray());
            //
            //StringBuilder translated_buffer = new StringBuilder();
            //foreach (var translated_item in translated_items)
            //{
            //    if (translated_item.Content.Length == 0)
            //    {
            //        translated_buffer.AppendLine("");
            //    }
            //    else
            //    {
            //        string indent = translated_item.IndentationDepth == 0 ? "" : new string(' ', translated_item.IndentationDepth * 4);
            //        translated_buffer.Append(indent).AppendLine(translated_item.Content);
            //    }
            //}

            //string translated_cs_actual = translated_buffer.ToString();
            if (translated_cs_expected != null)
            {
                if (translated_cs_expected != translated_cs_actual)
                {
                    storedFile = tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".cs", translated_cs_expected, translated_cs_actual);
                    int mismatchIndex = FindFirstMismatchIndex(translated_cs_expected, translated_cs_actual, out int mismatchLine, out int mismatchColumn, out char? mismatchCharA, out char? mismatchCharB);
                    string snippetE = GetMismatchedSnippet(translated_cs_expected, mismatchIndex, 100);
                    string snippetA = GetMismatchedSnippet(translated_cs_actual, mismatchIndex, 100);
                    failed_text = $"C# translation failed. See 'Output' for more information. {NewLineNormalized}Mismatch at line:{mismatchLine}, column:{mismatchColumn} (Index:{mismatchIndex}) {NewLineNormalized}E:'{snippetE}' {NewLineNormalized}A:'{snippetA}'. storedFile:" + storedFile;
                }
            }
            else
            {
                storedFile = tst.SaveExpectedActualFile(chainName, workItemName, chainName + ".cs", translated_cs_actual);
            }

            if (generated_vbs_expected == null)
            {
                //storedFile = tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".generated.vbs", generated_vbs_expected ?? "", generated_vbs_actual);
                storedFile = tst.SaveExpectedActualFile(chainName, workItemName, chainName + ".vbs", scriptContent);
            }

            if (!string.IsNullOrEmpty(failed_text))
            {
                Assert.Fail(failed_text);
            }

            return new TestScriptResponse(translated_cs_actual);
        }

        private static IOutermostScope FromXml(string xmlA)
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(IOutermostScope), OutermostScopeKnownTypes.AllKnownTypes);
            StringBuilder text_buffer = new StringBuilder();
            using (StringReader text_reader = new StringReader(xmlA))
            {
                using (System.Xml.XmlReader xReader = System.Xml.XmlReader.Create(text_reader))
                {
                    return (IOutermostScope)serializer.ReadObject(xReader);
                }
            }
        }

        private static string ToXml(IOutermostScope outermostBlock, Action<string> failed_handler)
        {
            string xmlA = ToXmlImpl(outermostBlock);
            var blockB = FromXml(xmlA);
            string xmlB = ToXmlImpl(blockB);

            if (xmlA != xmlB)
            {
                failed_handler("diff xml.");
            }


            return xmlB;
        }

        private static string ToXmlImpl(IOutermostScope blockSet)
        {
            DataContractSerializer serializer = new DataContractSerializer(typeof(IOutermostScope), OutermostScopeKnownTypes.AllKnownTypes);
            StringBuilder text_buffer = new StringBuilder();
            using (System.Xml.XmlWriter xWriter = System.Xml.XmlWriter.Create(text_buffer, new System.Xml.XmlWriterSettings()
            {
                Indent = true,
                OmitXmlDeclaration = true,
                NamespaceHandling = System.Xml.NamespaceHandling.OmitDuplicates

            }))
            {
                serializer.WriteObject(xWriter, blockSet);
                xWriter.Flush();
            }

            return text_buffer.ToString().NormalizeLineEndings();
        }
    }
    public enum ScriptUsageKind
    {
        Unknown,
        Connectivity,
        EBL,
        DialogGui, // model, named symboles, controls
        DialogWeb
    }

    public sealed class TestScriptResponse
    {
        public string TranslatedCsCode { get; }

        public TestScriptResponse(string translatedCsCode)
        {
            TranslatedCsCode = translatedCsCode ?? throw new ArgumentNullException(nameof(translatedCsCode));
        }
    }
}
