#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// Parity of the command line tool with cscript.exe where Windows is involved: the scripts that use the services of Windows (processes,
    /// clipboard, OLE DB, Outlook, WMI, registry, environment, event log, COM) are run with the command line tool and their output is compared
    /// with the output of cscript.exe - captured (samples\Windows, tests\Skrypton.Tests\VbsResources\ComParity_*.vbs) or of a cscript run of the
    /// test. The other parity scripts and demos run cross-platform in Skrypton.Tests (ComParityTests - issue #142).
    /// </summary>
    [TestClass]
    public sealed partial class CliParityTests
    {
        [TestMethod]
        [DataRow("Ado-StreamAndDatabase", "")] // OLE DB
        [DataRow("Edge-OpenUrl", "/dryrun")] // WshShell.Run
        [DataRow("Outlook-Mail", "/dryrun")] // Outlook
        [DataRow("Shell-Processes", "")] // processes (Exec, Run, WMI)
        public void DemoScript(string name, string argument)
        {
            string folder = Path.Combine(CliRunner.RepositoryRoot, "samples", "Windows");
            var arguments = new List<string> { "/culture:en-US", Path.Combine(folder, name + ".vbs") }; // the culture of the captured cscript output
            if (argument.Length > 0)
                arguments.Add(argument);
            CliRunner.Result result = CliRunner.Run(arguments);
            AssertSameOutput(name, File.ReadAllText(Path.Combine(folder, name + ".expected.txt")), result);
        }

        /// <summary>
        /// The parity scripts of the components that use the services of Windows - Skrypton.Tests runs them with test services
        /// </summary>
        [TestMethod]
        [DataRow("ComParity_WScriptShellScriptletWmiAdodbTypeNames")] // processes (WMI, ExpandEnvironmentStrings)
        [DataRow("ComParity_InternetExplorerHtmlFile")] // clipboard
        public void ComponentParityScript(string name)
        {
            string folder = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Tests", "VbsResources");
            string script = Path.Combine(Path.GetTempPath(), "skrypton_cli_" + Guid.NewGuid().ToString("N") + ".vbs");
            var environment = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            environment["SKRYPTON_PARITY_VAR"] = "parity-value";

            // the parity scripts write with Out (cscript: WScript.Echo)
            File.WriteAllText(script, "Sub Out(s) : WScript.Echo \"\" & s : End Sub\r\n" + File.ReadAllText(Path.Combine(folder, name + ".vbs")));
            bool usesClipboard = name.Contains("HtmlFile"); // the script uses the clipboard - one test at a time (they run in parallel)
            if (usesClipboard)
                System.Threading.Monitor.Enter(SharedMachineState.Clipboard);
            string? clipboard = usesClipboard ? new WindowsHostClipboardHostService().GetText("Text") : null;
            try
            {
                CliRunner.Result result = CliRunner.Run(new[] { "/culture:en-US", script }, environment);
                AssertSameOutput(name, File.ReadAllText(Path.Combine(folder, name + ".expected.txt")), result);
            }
            finally
            {
                File.Delete(script);
                if (clipboard != null)
                    new WindowsHostClipboardHostService().SetText("Text", clipboard);
                if (usesClipboard)
                    System.Threading.Monitor.Exit(SharedMachineState.Clipboard);
            }
        }

        /// <summary>
        /// WMI Win32_TimeZone (issue #101): its values depend on the time zone of the machine, so the script runs with cscript (and the real
        /// WMI) and with the command line tool here - the outputs must be the same
        /// </summary>
        [TestMethod]
        public void WmiTimeZoneIsTheOneOfWindows()
        {
            string script = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Windows.Tests", "WmiTimeZone.vbs");
            var cscript = new System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cscript.exe"), "//NoLogo \"" + script + "\"")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            string expected;
            using (System.Diagnostics.Process process = System.Diagnostics.Process.Start(cscript)!)
            {
                expected = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                Assert.AreEqual(0, process.ExitCode, "cscript: " + expected + process.StandardError.ReadToEnd());
            }
            // cscript formats the date with the culture of the user, as does the command line tool with /culture
            CliRunner.Result result = CliRunner.Run(new[] { "/culture:" + System.Globalization.CultureInfo.CurrentCulture.Name, script });
            AssertSameOutput("WmiTimeZone", expected, result);
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
