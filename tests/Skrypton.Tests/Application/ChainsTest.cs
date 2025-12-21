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
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;

namespace Skrypton.Tests.Application
{
    // D:\zapechene.2015\VBScript.Parse\LuboVBParser1\TestResources
    [TestClass]
    public class ChainsTest : TestBase
    {
        [TestMethod, MyMemberData(nameof(ChainNames))]
        public void Chains(string chainName)
        {
            TestCncInChain(chainName);
        }

        public static object[][] ChainNames
        {
            get
            {
                List<string> names = new List<string>();
                ///names.Add("aaaaaa");

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
                foreach (var arg0 in names)
                {
                    result.Add(new object[] { arg0 });
                }

                return result.ToArray();
            }
        }


        public void TestCncInChain(string chainName)
        {
            string x_ressource_name = chainName;
            string scriptContent = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".vbs");
            string generated_vbs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".generated.vbs");
            string translated_cs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".cstxt");
            string xml_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + ".xml");

            List<string> externalDependencies = new List<string>();
            if (scriptContent.Contains("hlContext"))
                externalDependencies.Add("hlContext"); // EBL
            else
                externalDependencies.Add("session"); // Connectivity IN/OUT

            //Console.WriteLine("parsing...");
            var parsed_items = Skrypton.LegacyParser.Parser.Parse(TestCulture, scriptContent);

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
                SaveExpectedActualFiles(chainName, workItemName, chainName + ".generated.vbs", generated_vbs_expected, generated_vbs_actual);
                failed_text = "VBS generation failed. See 'Output' for more information.";
            }

            var outermostBlock = Skrypton.LegacyParser.Parser.ParseToOutermostScope(TestCulture, scriptContent);
            var xml_actual = ToXml(outermostBlock, x => failed_text = x);

            if (xml_expected != xml_actual)
            {
                SaveExpectedActualFiles(chainName, workItemName, chainName + ".xml", xml_expected, xml_actual);
                failed_text = "Xml generation failed. See 'Output' for more information.";
            }


            Console.WriteLine("translating...");
            IEnumerable<TranslatedStatement> translated_items = Skrypton.CSharpWriter.DefaultTranslator.Translate(TestCulture, scriptContent, externalDependencies.ToArray());

            StringBuilder translated_buffer = new StringBuilder();
            foreach (var translated_item in translated_items)
            {
                if (translated_item.Content.Length == 0)
                {
                    translated_buffer.AppendLine("");
                }
                else
                {
                    string indent = translated_item.IndentationDepth == 0 ? "" : new string(' ', translated_item.IndentationDepth * 4);
                    translated_buffer.Append(indent).AppendLine(translated_item.Content);
                }
            }

            string translated_cs_actual = translated_buffer.ToString();
            if (translated_cs_expected != translated_cs_actual)
            {
                SaveExpectedActualFiles(chainName, workItemName, chainName + ".cs.txt", translated_cs_expected, translated_cs_actual);
                failed_text = "C# translation failed. See 'Output' for more information.";
            }

            if (!string.IsNullOrEmpty(failed_text))
            {
                Assert.Fail(failed_text);
            }
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

        private string ToXml(IOutermostScope outermostBlock, Action<string> failed_handler)
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
}
