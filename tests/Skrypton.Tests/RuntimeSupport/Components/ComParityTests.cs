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
        [TestMethod] public void ComParity_BuiltInFunctions() => RunParityScript();
        [TestMethod] public void ComParity_ErrObject() => RunParityScript();
        [TestMethod] public void ComParity_UndeclaredVariables() => RunParityScript();
        [TestMethod] public void ComParity_ByteFunctions() => RunParityScript();

        /// <summary>
        /// The parity scripts run with VBScript's rule for undeclared variables (as cscript and the command line tools - issue #86)
        /// </summary>
        private ScriptControlConfiguration VBScriptRuleConfiguration() => new TestScriptControlConfiguration(this, false, [], [], undeclaredVariablesAreErrors: false);

        /// <summary>
        /// Skrypton's default (strict) rule: an undeclared variable is a translation error, also without Option Explicit (issue #86)
        /// </summary>
        [TestMethod]
        public void UndeclaredVariablesAreErrorsByDefault()
        {
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            Exception exception = Assert.Throws<Exception>(() => scriptengine.ExecuteStatement("x = 1\r\n"));
            StringAssert.Contains(exception.ToString(), "SKY101");

            IScriptControl vbscriptRule = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), VBScriptRuleConfiguration());
            vbscriptRule.ExecuteStatement("x = 1\r\nx = x + 1\r\n");
        }

        /// <summary>
        /// Unhandled errors (ComParity_UnhandledErrors.cases.txt): each script fails and the error that the script control reports - "(line)
        /// column) source: description" - must be the one that Windows cscript.exe reports for it (issues #84, #85, #86, #87)
        /// </summary>
        [TestMethod]
        public void ComParity_UnhandledErrors()
        {
            var failures = new List<string>();
            foreach ((string script, string expected) in UnhandledErrorCases())
            {
                IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), VBScriptRuleConfiguration());
                try
                {
                    scriptengine.ExecuteStatement(script);
                    failures.Add($"{script}{Environment.NewLine}  expected: {expected}{Environment.NewLine}  no error");
                    continue;
                }
                catch (Exception)
                {
                    // the error is in scriptengine.Error
                }
                Error error = scriptengine.Error!;
                // Error.Column is 0-based (as MSScriptControl's), cscript reports it 1-based (issue #87)
                string actual = $"({error.Line}, {error.Column + 1}) {error.Source}: {error.Description}";
                if (actual != expected)
                    failures.Add($"{script}{Environment.NewLine}  expected: {expected}{Environment.NewLine}  actual:   {actual}");
            }
            Assert.IsEmpty(failures, string.Join(Environment.NewLine, failures));
        }

        /// <summary>
        /// The line and the column of the failing statement in a script with more lines, as cscript reports them: "(4, 13)" - a ':' and a
        /// comment on the line do not count as statements (issue #87)
        /// </summary>
        [TestMethod]
        [DataRow("Dim a\r\na = 1\r\nSub Go(o)\r\n    a = 2 : o.Run \"x\" ' comment: o.Fail\r\nEnd Sub\r\nGo Nothing\r\n", 4, 13, "a = 2 : o.Run \"x\" ' comment: o.Fail")]
        [DataRow("x = \"a:b\" : y = 1 / 0\r\n", 1, 13, "x = \"a:b\" : y = 1 / 0")]
        public void ErrorReportsTheLineAndColumnOfTheFailingStatement(string script, int line, int column, string text)
        {
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), VBScriptRuleConfiguration());
            Assert.Throws<Exception>(() => scriptengine.ExecuteStatement(script));
            Error error = scriptengine.Error!;
            Assert.AreEqual(line, error.Line);
            Assert.AreEqual(column - 1, error.Column); // 0-based
            Assert.AreEqual(text, error.Text);
        }

        /// <summary>
        /// The cases of ComParity_UnhandledErrors.cases.txt: a script (one line) and the error that cscript reports for it
        /// </summary>
        internal static IEnumerable<(string Script, string Expected)> UnhandledErrorCases()
        {
            string[] lines = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources.ComParity_UnhandledErrors.cases.txt")
                .Replace("\r\n", "\n").Split('\n')
                .Where(line => (line.Length > 0) && !line.StartsWith("#", StringComparison.Ordinal))
                .ToArray();
            for (int i = 0; i + 1 < lines.Length; i += 2)
                yield return (lines[i], lines[i + 1]);
        }

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
            // The expected output was captured with cscript on Windows with the culture en-US (numbers, currency, month names)
            TestCulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");
            var output = new Output();
            IScriptControl scriptengine = CreateScriptControlClass(
                // WMI (winmgmts:) lists the processes of the host's process service (cscript lists the real ones; the script only checks that there are some)
                // htmlfile's clipboardData uses an in-memory clipboard (cscript's htmlfile does not change the user's clipboard either)
                CreateRuntimeHost(CreateTestHostServices(s => s
                    .RegisterHostService<IHostProcessControlHostService>(() => new TestHostProcessControlHostService())
                    .RegisterHostService<IHostClipboardHostService>(() => new TestClipboardHostService())
                    .RegisterHostService<IHostOutlookHostService>(() => new TestOutlookHostService()))),
                VBScriptRuleConfiguration()
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
