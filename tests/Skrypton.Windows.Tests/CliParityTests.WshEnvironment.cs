#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// WshShell.Environment (issue #107): the script runs with cscript and with the command line tool here (both with the environment of the
        /// test process and the machine's registry) - the outputs must be the same
        /// </summary>
        [TestMethod]
        public void WshEnvironmentIsTheOneOfWindows()
        {
            string script = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Windows.Tests", "WshEnvironment.vbs");
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
            AssertSameOutput("WshEnvironment", expected, CliRunner.RunProcess(new[] { script }));
        }
    }
}
