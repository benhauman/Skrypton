#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// WshShell.LogEvent (issues #116, #139): the script runs with cscript and with the command line tool - both write the same entries to the
        /// Application event log: the source WSH, the type as the event id with the severity bits of WSH (the qualifiers, without them the event
        /// log can not find the message) and the level (0 and the audit types have level 0, 1 is error = 2, 2 warning = 3, 4 information = 4)
        /// </summary>
        [TestMethod]
        public void LogEventWritesToTheApplicationLogAsWsh()
        {
            string marker = "skrypton-logevent-" + Guid.NewGuid().ToString("N");
            string script = Path.Combine(Path.GetTempPath(), marker + ".vbs");
            string query = Path.Combine(Path.GetTempPath(), marker + ".ps1");
            File.WriteAllText(script,
                "Dim sh : Set sh = CreateObject(\"WScript.Shell\")\r\n" +
                "Dim t\r\n" +
                "For Each t In Array(0, 1, 2, 4, 8, 16)\r\n" +
                "    sh.LogEvent t, WScript.Arguments(0) & \" \" & t\r\n" +
                "Next\r\n" +
                "WScript.Echo \"\" & sh.LogEvent(4, WScript.Arguments(0) & \" result\")\r\n");
            // the message is the one that the event log finds with the message file of WSH (the description of the event)
            File.WriteAllText(query,
                "param([string]$Marker)\r\n" +
                "Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'WSH' } -MaxEvents 50 |\r\n" +
                "    Where-Object { $_.Message -like \"*$Marker*\" } |\r\n" +
                "    Sort-Object RecordId |\r\n" +
                "    ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Id, $_.Qualifiers, $_.Level, ($_.Message -replace [regex]::Escape($Marker), 'M') }\r\n");
            try
            {
                string cscriptMarker = marker + "-cscript";
                var cscript = new System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cscript.exe"), "//NoLogo \"" + script + "\" " + cscriptMarker)
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                };
                string cscriptOutput;
                using (System.Diagnostics.Process process = System.Diagnostics.Process.Start(cscript)!)
                {
                    cscriptOutput = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                }

                string skryptonMarker = marker + "-skrypton";
                CliRunner.Result result = CliRunner.Run(new[] { script, skryptonMarker });
                Assert.AreEqual("", result.StandardError);
                Assert.AreEqual(cscriptOutput.Trim(), result.StandardOutput.Trim());
                Assert.AreEqual("True", result.StandardOutput.Trim());

                string[] expected = { "0|0|0|M 0", "1|49152|2|M 1", "2|32768|3|M 2", "4|16384|4|M 4", "8|0|0|M 8", "16|49152|0|M 16", "4|16384|4|M result" };
                CollectionAssert.AreEqual(expected, Entries(query, cscriptMarker), "cscript");
                CollectionAssert.AreEqual(expected, Entries(query, skryptonMarker), "the command line tool");
            }
            finally
            {
                File.Delete(script);
                File.Delete(query);
            }
        }

        private static string[] Entries(string query, string marker)
        {
            var powershell = new System.Diagnostics.ProcessStartInfo("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -File \"" + query + "\" " + marker)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            };
            using (System.Diagnostics.Process process = System.Diagnostics.Process.Start(powershell)!)
            {
                string entries = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                return CliRunner.Lines(entries);
            }
        }
    }
}
