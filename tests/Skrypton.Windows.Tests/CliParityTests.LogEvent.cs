#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// WshShell.LogEvent (issue #116): the command line tool writes the entries to the Application event log as Windows Script Host does - the
        /// source WSH, the type as the event id and the level (0 and the audit types have level 0, 1 is error = 2, 2 warning = 3, 4 information = 4)
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
                "    sh.LogEvent t, \"" + marker + " \" & t\r\n" +
                "Next\r\n" +
                "WScript.Echo \"\" & sh.LogEvent(4, \"" + marker + " result\")\r\n");
            File.WriteAllText(query,
                "Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'WSH' } -MaxEvents 50 |\r\n" +
                "    Where-Object { (($_.Properties | ForEach-Object { [string]$_.Value }) -join ' ') -like '*" + marker + "*' } |\r\n" +
                "    Sort-Object RecordId |\r\n" +
                "    ForEach-Object { '{0}|{1}|{2}' -f $_.Id, $_.Level, ((($_.Properties | ForEach-Object { [string]$_.Value }) -join ' ') -replace '" + marker + "', 'M') }\r\n");
            try
            {
                CliRunner.Result result = CliRunner.Run(new[] { script });
                Assert.AreEqual("", result.StandardError);
                Assert.AreEqual("True", result.StandardOutput.Trim());

                var powershell = new System.Diagnostics.ProcessStartInfo("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -File \"" + query + "\"")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                };
                string entries;
                using (System.Diagnostics.Process process = System.Diagnostics.Process.Start(powershell)!)
                {
                    entries = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                }
                CollectionAssert.AreEqual(
                    new[] { "0|0|M 0", "1|2|M 1", "2|3|M 2", "4|4|M 4", "8|0|M 8", "16|0|M 16", "4|4|M result" },
                    CliRunner.Lines(entries));
            }
            finally
            {
                File.Delete(script);
                File.Delete(query);
            }
        }
    }
}
