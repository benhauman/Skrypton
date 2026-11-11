#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// Parity of the command line tool with cscript.exe: the demo scripts (samples\Windows) and the parity scripts of the components
    /// (tests\Skrypton.Tests\VbsResources\ComParity_*.vbs) are run with the command line tool and their output is compared with the output that
    /// was captured with cscript.exe (and Word, Excel, Outlook, Internet Explorer) on Windows
    /// </summary>
    [TestClass]
    public sealed class CliParityTests
    {
        [TestMethod]
        [DataRow("Ado-StreamAndDatabase", "")]
        [DataRow("Edge-OpenUrl", "/dryrun")]
        [DataRow("Excel-Report", "")]
        [DataRow("FileSystem-Inventory", "")]
        [DataRow("Outlook-Mail", "/dryrun")]
        [DataRow("Shell-Processes", "")]
        [DataRow("Word-Letter", "")]
        public void DemoScript(string name, string argument)
        {
            string folder = Path.Combine(CliRunner.RepositoryRoot, "samples", "Windows");
            var arguments = new List<string> { "/culture:en-US", Path.Combine(folder, name + ".vbs") }; // the culture of the captured cscript output
            if (argument.Length > 0)
                arguments.Add(argument);
            CliRunner.Result result = CliRunner.Run(arguments);
            AssertSameOutput(name, File.ReadAllText(Path.Combine(folder, name + ".expected.txt")), result);
        }

        [TestMethod]
        [DataRow("ComParity_ScriptingDictionary")]
        [DataRow("ComParity_VBScriptRegExp")]
        [DataRow("ComParity_Msxml2DOMDocument")]
        [DataRow("ComParity_ScriptingFileSystemObject")]
        [DataRow("ComParity_WScriptShellScriptletWmiAdodbTypeNames")]
        [DataRow("ComParity_InternetExplorerHtmlFile")]
        [DataRow("ComParity_WordApplication")]
        [DataRow("ComParity_ExcelApplication")]
        [DataRow("ComParity_BuiltInFunctions")]
        [DataRow("ComParity_ErrObject")]
        public void ComponentParityScript(string name)
        {
            string folder = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Tests", "VbsResources");
            string script = Path.Combine(Path.GetTempPath(), "skrypton_cli_" + Guid.NewGuid().ToString("N") + ".vbs");
            var environment = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            environment["SKRYPTON_PARITY_VAR"] = "parity-value";
            var templates = new List<string>();
            foreach (string extension in new[] { "docx", "xlsx" })
            {
                string template = Path.Combine(folder, name + ".template." + extension);
                if (!File.Exists(template))
                    continue;
                string copy = Path.Combine(Path.GetTempPath(), "skrypton_cli_template_" + Guid.NewGuid().ToString("N") + "." + extension);
                File.Copy(template, copy);
                templates.Add(copy);
                environment["SKRYPTON_PARITY_" + extension.ToUpperInvariant()] = copy;
            }

            // the parity scripts write with Out (cscript: WScript.Echo)
            File.WriteAllText(script, "Sub Out(s) : WScript.Echo \"\" & s : End Sub\r\n" + File.ReadAllText(Path.Combine(folder, name + ".vbs")));
            string? clipboard = name.Contains("HtmlFile") ? new WindowsHostClipboardHostService().GetText("Text") : null; // the script uses the clipboard
            try
            {
                CliRunner.Result result = CliRunner.Run(new[] { "/culture:en-US", script }, environment);
                AssertSameOutput(name, File.ReadAllText(Path.Combine(folder, name + ".expected.txt")), result);
            }
            finally
            {
                File.Delete(script);
                foreach (string template in templates)
                    File.Delete(template);
                if (clipboard != null)
                    new WindowsHostClipboardHostService().SetText("Text", clipboard);
            }
        }


        /// <summary>
        /// Unhandled errors (tests\Skrypton.Tests\VbsResources\ComParity_UnhandledErrors.cases.txt): the command line tool must report each
        /// error as cscript does - "script.vbs(line, column) source: description" (compared without the file and the column) - and end with 1
        /// </summary>
        [TestMethod]
        public void UnhandledErrorsAreReportedAsCscriptReportsThem()
        {
            string[] lines = File.ReadAllLines(Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Tests", "VbsResources", "ComParity_UnhandledErrors.cases.txt"))
                .Where(line => (line.Length > 0) && !line.StartsWith("#", StringComparison.Ordinal))
                .ToArray();
            string script = Path.Combine(Path.GetTempPath(), "skrypton_cli_" + Guid.NewGuid().ToString("N") + ".vbs");
            var failures = new List<string>();
            try
            {
                for (int i = 0; i + 1 < lines.Length; i += 2)
                {
                    File.WriteAllText(script, lines[i] + "\r\n");
                    CliRunner.Result result = CliRunner.Run(new[] { "/culture:en-US", script });
                    string reported = result.StandardError.Trim();
                    // "C:\..\x.vbs(1, 41) Microsoft VBScript runtime error: .." -> "(1) Microsoft VBScript runtime error: .."
                    string actual = System.Text.RegularExpressions.Regex.Replace(reported, @"^.*?\((\d+), \d+\) ", "($1) ");
                    if ((actual != lines[i + 1]) || (result.ExitCode != 1))
                        failures.Add($"{lines[i]}{Environment.NewLine}  expected: {lines[i + 1]}{Environment.NewLine}  actual:   {actual} (exit code {result.ExitCode})");
                }
            }
            finally
            {
                File.Delete(script);
            }
            Assert.IsEmpty(failures, string.Join(Environment.NewLine, failures));
        }

        private static void AssertSameOutput(string name, string expected, CliRunner.Result result)
        {
            string[] expectedLines = CliRunner.Lines(expected);
            string[] actualLines = CliRunner.Lines(result.StandardOutput);
            string details = $"{name}: exit code {result.ExitCode}{Environment.NewLine}standard error:{Environment.NewLine}{result.StandardError}{Environment.NewLine}output:{Environment.NewLine}{result.StandardOutput}";
            Assert.AreEqual("", result.StandardError, details);
            for (int i = 0; i < Math.Min(expectedLines.Length, actualLines.Length); i++)
                Assert.AreEqual(expectedLines[i], actualLines[i], $"{name}: line {i + 1} differs from cscript{Environment.NewLine}{details}");
            Assert.HasCount(expectedLines.Length, actualLines, details);
            Assert.AreEqual(0, result.ExitCode, details);
        }
    }
}
