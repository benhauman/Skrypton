#nullable enable
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// Registered COM classes (issue #103): with /com the command line tool creates them as cscript does - methods, properties (get and set),
        /// For Each, a script object passed to the COM object and called back, error 429 for an unregistered ProgID. Without /com the ProgID is
        /// error 429.
        /// </summary>
        [TestMethod]
        public void ComActivationCreatesRegisteredComClasses()
        {
            string script = Path.Combine(CliRunner.RepositoryRoot, "tests", "Skrypton.Windows.Tests", "ComActivation.vbs");
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
            AssertSameOutput("ComActivation", expected, CliRunner.Run(new[] { "/com:WScript.Network,System.Collections.ArrayList", script }));

            CliRunner.Result withoutOptIn = CliRunner.Run(new[] { script });
            Assert.AreEqual(1, withoutOptIn.ExitCode);
            StringAssert.Contains(withoutOptIn.StandardError, "Microsoft VBScript runtime error: ActiveX component can't create object: 'WScript.Network'");
        }
    }
}
