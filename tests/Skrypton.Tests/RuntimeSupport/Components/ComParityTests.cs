#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// Black-box parity tests for the built-in COM components: each VbsResources\ComParity_*.vbs script is run through Skrypton and its output is
    /// compared line by line with ComParity_*.expected.txt, which was captured by running the same script with Windows cscript.exe and the real
    /// COM components (scrrun.dll, vbscript.dll, msxml, wshom, wbem, ADO). The scripts write their results with "Out", which both sides convert to
    /// text ("" &amp; value).
    /// </summary>
    [TestClass]
    public sealed class ComParityTests : TestBase
    {
        public sealed class Output
        {
            public List<string> Lines { get; } = new List<string>();
            public void Write(object? value) => Lines.Add(Convert.ToString(value) ?? "");
        }

        [TestMethod] public void ComParity_ScriptingDictionary() => RunParityScript();
        [TestMethod] public void ComParity_VBScriptRegExp() => RunParityScript();
        [TestMethod] public void ComParity_Msxml2DOMDocument() => RunParityScript();
        [TestMethod] public void ComParity_ScriptingFileSystemObject() => RunParityScript();
        [TestMethod] public void ComParity_WScriptShellScriptletWmiAdodbTypeNames() => RunParityScript();
        [TestMethod] public void ComParity_InternetExplorerHtmlFile() => RunParityScript();
        [TestMethod] public void ComParity_WordApplication() => RunParityScript();
        [TestMethod] public void ComParity_ExcelApplication() => RunParityScript();
        [TestMethod] public void ComParity_OutlookApplication() => RunParityScript();

        private void RunParityScript()
        {
            string script = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + ".vbs");
            string expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + TestName + ".expected.txt");

            // The script reads this variable with WScript.Shell.ExpandEnvironmentStrings (it was set the same way for cscript)
            Environment.SetEnvironmentVariable("SKRYPTON_PARITY_VAR", "parity-value");

            // An Office template of the script (ComParity_*.template.docx / .xlsx, created with Word / Excel) is passed as a file, its path in
            // SKRYPTON_PARITY_DOCX / SKRYPTON_PARITY_XLSX (as for cscript)
            string? templateFile = null;
            foreach (string extension in new[] { "docx", "xlsx" })
            {
                using System.IO.Stream? template = typeof(ComParityTests).Assembly.GetManifestResourceStream("Skrypton.Tests.VbsResources." + TestName + ".template." + extension);
                if (template == null)
                    continue;
                templateFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "skrypton_parity_template_" + Guid.NewGuid().ToString("N") + "." + extension);
                using (System.IO.FileStream file = System.IO.File.Create(templateFile))
                    template.CopyTo(file);
                Environment.SetEnvironmentVariable("SKRYPTON_PARITY_" + extension.ToUpperInvariant(), templateFile);
            }
            try
            {
                RunParityScript(script, expected);
            }
            finally
            {
                if (templateFile != null)
                    System.IO.File.Delete(templateFile);
            }
        }

        private void RunParityScript(string script, string expected)
        {

            var output = new Output();
            IScriptControl scriptengine = CreateScriptControlClass(
                // WMI (winmgmts:) lists the processes of the host's process service (cscript lists the real ones; the script only checks that there are some)
                // htmlfile's clipboardData uses an in-memory clipboard (cscript's htmlfile does not change the user's clipboard either)
                CreateRuntimeHost(CreateTestHostServices(s => s
                    .RegisterHostService<IHostProcessControlHostService>(() => new TestHostProcessControlHostService())
                    .RegisterHostService<IHostClipboardHostService>(() => new TestClipboardHostService())
                    .RegisterHostService<IHostOutlookHostService>(() => new TestOutlookHostService()))),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("output", output, false);
            scriptengine.ExecuteStatement("Sub Out(s)\r\n  output.Write \"\" & s\r\nEnd Sub\r\n" + script.Replace("\r\n", "\n").Replace("\n", "\r\n"));

            string[] expectedLines = expected.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
            string[] actualLines = string.Join("\n", output.Lines).Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
            for (int i = 0; i < Math.Min(expectedLines.Length, actualLines.Length); i++)
                Assert.AreEqual(expectedLines[i], actualLines[i], $"{TestName}: line {i + 1} differs from Windows VBScript");
            Assert.HasCount(expectedLines.Length, actualLines, $"{TestName}: number of output lines (expected:{Environment.NewLine}{string.Join(Environment.NewLine, expectedLines)}{Environment.NewLine}actual:{Environment.NewLine}{string.Join(Environment.NewLine, actualLines)})");
        }
    }
}
