using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;
using Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Skrypton.Tests.Application
{
    // D:\zapechene.2015\VBScript.Parse\LuboVBParser1\TestResources
    [TestClass]
    public sealed class ChainsTest : TestBase
    {
        [TestMethod, MyMemberData(nameof(ChainNames))]
        public void Chains(string chainName, ScriptUsageKind scriptUsage)
        {
            TestScriptChain(this, chainName, scriptUsage);
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
                    result.Add(new object[] { chainName, scriptUsage });
                }

                return result.ToArray();
            }
        }
        internal static void TestScriptChain(TestBase tst, string chainName, ScriptUsageKind scrUsage, Dictionary<string, object> externalRefs = null)
        {
            string x_ressource_name = chainName;
            string scriptContent = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".vbs");
            string generated_vbs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".generated.vbs");
            string translated_cs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + CSFileExtension);
            string xml_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".xml");

            NonNullImmutableList<string> externalDependencies = new NonNullImmutableList<string>();
            if (scrUsage == ScriptUsageKind.EBL)//(scriptContent.Contains("hlContext"))
                externalDependencies = externalDependencies.Add("hlContext"); // EBL
            if (scrUsage == ScriptUsageKind.Connectivity)
                externalDependencies = externalDependencies.Add("session"); // Connectivity IN/OUT
            if (externalRefs != null)
            {
                foreach (string externalRefName in externalRefs.Keys)
                {
                    externalDependencies = externalDependencies.Add(externalRefName);
                }
            }

            //Console.WriteLine("parsing...");
            var parsed_items = Skrypton.LegacyParser.Parser.Parse(tst.TestCulture, scriptContent);

            StringBuilder parsed_output = new StringBuilder();
            ISourceIndentHandler parsed_intender = new Skrypton.LegacyParser.CodeBlocks.SourceRendering.SourceIndentHandler();
            foreach (ICodeBlock parsed_item in parsed_items)
            {
                parsed_output.AppendLine(parsed_item.GenerateBaseSource(parsed_intender));
            }

            string workItemName = "Script";// TestContext.TestName;
            string generated_vbs_actual = parsed_output.ToString();

            string failed_text = null;

            if (generated_vbs_expected != generated_vbs_actual)
            {
                tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".generated.vbs", generated_vbs_expected, generated_vbs_actual);
                failed_text = "VBS generation failed. See 'Output' for more information.";
            }

            var outermostBlock = Skrypton.LegacyParser.Parser.ParseToOutermostScope(tst.TestCulture, scriptContent);
            var xml_actual = ToXml(outermostBlock, x => failed_text = x);

            if (xml_expected != xml_actual)
            {
                tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".xml", xml_expected, xml_actual);
                failed_text = "Xml generation failed. See 'Output' for more information.";
            }


            Console.WriteLine("translating...");
            var csLines = DefaultCSharpTranslation.GetTranslatedStatements(tst.TestCulture, scriptContent, externalDependencies);
            string translated_cs_actual = string.Join("\r\n", csLines);

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
            if (translated_cs_expected != translated_cs_actual)
            {
                tst.SaveExpectedActualFiles(chainName, workItemName, chainName + ".cs.txt", translated_cs_expected, translated_cs_actual);
                int mismatchIndex = FindFirstMismatchIndex(translated_cs_expected, translated_cs_actual, out int mismatchLine, out int mismatchColumn);
                string snippetE = GetMismatchedSnippet(translated_cs_expected, mismatchIndex, 100);
                string snippetA = GetMismatchedSnippet(translated_cs_actual, mismatchIndex, 100);
                failed_text = $"C# translation failed. See 'Output' for more information. \r\nMismatch at line:{mismatchLine}, column:{mismatchColumn} (Index:{mismatchIndex}) \r\nE:'{snippetE}' \r\nA:'{snippetA}'";
            }

            if (!string.IsNullOrEmpty(failed_text))
            {
                Assert.Fail(failed_text);
            }
        }
        private static int FindFirstMismatchIndex(string a, string b, out int line, out int column)
        {
            line = 1;
            column = 1;

            int minLength = Math.Min(a.Length, b.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (a[i] != b[i])
                    return i;
                if (a[i] == '\n') // handle windows and unix line endings
                {
                    line++;
                    column = 1;
                }
                else if (a[i] != '\r') // ignore carriage return
                {
                    column++;
                }
            }
            if (a.Length != b.Length)
                return minLength;
            return -1; // no mismatch
        }
        private static string GetMismatchedSnippet(string s, int startIndex, int maxLength)
        {
            if (startIndex > s.Length)
                return "";
            int endOfLine = s.IndexOfAny(new char[] { '\r', '\n' }, startIndex);
            if (endOfLine == -1)
                endOfLine = s.Length;

            //int remaining  = s.Length - startIndex;
            int take = Math.Min(maxLength, endOfLine - startIndex);
            return s.Substring(startIndex, take);
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

            return text_buffer.ToString();
        }


    }
    public enum ScriptUsageKind
    {
        Unknown,
        Connectivity,
        EBL,
        DialogGui, // model, named symboles, controls
        DiagloWeb
    }
}
