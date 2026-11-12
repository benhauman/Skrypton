#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// WScript.Echo (issue #119): the script runs with cscript and with the command line tool (with the culture of the user, as cscript) - the
        /// outputs must be the same
        /// </summary>
        [TestMethod]
        public void EchoWritesValuesAsCscript()
        {
            string script = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Windows.Tests", "Echo.vbs");
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
            AssertSameOutput("Echo", expected, CliRunner.Run(new[] { "/culture:" + System.Globalization.CultureInfo.CurrentCulture.Name, script }));
        }
    }
}
