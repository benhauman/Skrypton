#nullable enable
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// The command line tool in a 32-bit process (issue #147): a COM call with more than one argument (Fields(key) = value puts the value with
        /// the key - two VARIANTs) as 32-bit cscript does. The VARIANTs are 16 bytes there, not 24 - with 24 the second argument was read at the
        /// wrong place. The .NET Framework tool is AnyCPU: a copy of it is marked 32-bit (the 32BITREQUIRED flag of its CLR header), so no 32-bit
        /// build is needed. Only here (Skrypton.Windows.NetFx.Tests): the .NET tool has a 64-bit native launcher.
        /// </summary>
        [TestMethod]
        public void ComCallsWithSeveralArgumentsWorkInA32BitProcess()
        {
            string folder = Path.Combine(Path.GetTempPath(), "skrypton-cli32-" + Guid.NewGuid().ToString("N"));
            string script = Path.Combine(folder, "cdo-fields.vbs");
            Directory.CreateDirectory(folder);
            try
            {
                string cli32 = CopyAs32Bit(CliRunner.CliPath, folder);
                File.WriteAllText(script,
                    "Const cdoSendUsingMethod = \"http://schemas.microsoft.com/cdo/configuration/sendusing\"\r\n" +
                    "Const cdoSMTPServer = \"http://schemas.microsoft.com/cdo/configuration/smtpserver\"\r\n" +
                    "Const cdoSMTPServerPort = \"http://schemas.microsoft.com/cdo/configuration/smtpserverport\"\r\n" +
                    "Dim oCfg : Set oCfg = CreateObject(\"CDO.Configuration\")\r\n" +
                    "Dim oCfgFields : Set oCfgFields = oCfg.Fields\r\n" +
                    "oCfgFields(cdoSendUsingMethod) = 2\r\n" +
                    "oCfgFields(cdoSMTPServer) = \"127.0.0.1\"\r\n" +
                    "oCfgFields.Item(cdoSMTPServerPort).Value = 2525\r\n" +
                    "oCfgFields.Update\r\n" +
                    "WScript.Echo TypeName(oCfgFields) & \" \" & oCfgFields(cdoSendUsingMethod) & \" \" & oCfgFields.Item(cdoSMTPServer).Value & \" \" & oCfg.Fields(cdoSMTPServer) & \" \" & oCfgFields(cdoSMTPServerPort)\r\n");

                string expected = Run(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "cscript.exe"), "//NoLogo \"" + script + "\"", out int cscriptExitCode, out string cscriptError);
                Assert.AreEqual(0, cscriptExitCode, "32-bit cscript: " + expected + cscriptError);

                string actual = Run(cli32, "/com:CDO.Configuration \"" + script + "\"", out int exitCode, out string error);
                AssertSameOutput("ComCallsWithSeveralArgumentsWorkInA32BitProcess", expected, new CliRunner.Result(exitCode, actual, error));
            }
            finally
            {
                Directory.Delete(folder, recursive: true);
            }
        }

        /// <summary>
        /// Copy the tool (and the assemblies next to it) to the folder and mark the copy 32-bit: COR20 header flag 32BITREQUIRED (0x2)
        /// </summary>
        private static string CopyAs32Bit(string cliPath, string folder)
        {
            foreach (string file in Directory.GetFiles(Path.GetDirectoryName(cliPath)!))
                File.Copy(file, Path.Combine(folder, Path.GetFileName(file)));
            string copy = Path.Combine(folder, Path.GetFileName(cliPath));
            int flagsOffset;
            using (var stream = File.OpenRead(copy))
            using (var reader = new PEReader(stream))
                flagsOffset = reader.PEHeaders.CorHeaderStartOffset + 16; // cb (4), runtime version (2 + 2), metadata directory (8), flags
            using (var stream = new FileStream(copy, FileMode.Open, FileAccess.ReadWrite))
            {
                var flags = new byte[4];
                stream.Position = flagsOffset;
                Assert.AreEqual(4, stream.Read(flags, 0, 4));
                flags[0] |= 0x2;
                stream.Position = flagsOffset;
                stream.Write(flags, 0, 4);
            }
            return copy;
        }

        private static string Run(string fileName, string arguments, out int exitCode, out string standardError)
        {
            var startInfo = new ProcessStartInfo(fileName, arguments)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            using (Process process = Process.Start(startInfo)!)
            {
                var error = process.StandardError.ReadToEndAsync();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                exitCode = process.ExitCode;
                standardError = error.Result;
                return output;
            }
        }
    }
}
