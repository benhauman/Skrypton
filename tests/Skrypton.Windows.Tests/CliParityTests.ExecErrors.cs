#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// WshShell.Exec and Run of a missing program (issue #165): the script runs with cscript and with the command line tool (the Windows errors) - the outputs must be the
        /// same (the HRESULT, the source and the description of WSH)
        /// </summary>
        [TestMethod]
        public void ExecAndRunOfAMissingProgramRaiseTheErrorsOfWsh()
        {
            string script = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Windows.Tests", "ExecErrors.vbs");
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
            AssertSameOutput("ExecErrors", expected, CliRunner.Run(new[] { script }));
        }
    }
}
