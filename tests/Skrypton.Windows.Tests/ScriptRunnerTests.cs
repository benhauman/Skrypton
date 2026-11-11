#nullable enable
using System;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// ScriptonScriptRunner (in the process) and the command line tool: WScript, exit codes, errors as cscript reports them, help
    /// </summary>
    [TestClass]
    public sealed class ScriptRunnerTests
    {
        private static (int ExitCode, string Output, string Error) Run(string script, string[]? arguments = null, string input = "")
        {
            var output = new StringWriter();
            var error = new StringWriter();
            var runner = new ScriptonScriptRunner(WindowsRuntimeHost.CreateDefault(), output, error, new StringReader(input), CultureInfo.GetCultureInfo("en-US"));
            int exitCode = runner.Run(script, @"C:\scripts\test.vbs", arguments ?? Array.Empty<string>());
            return (exitCode, output.ToString(), error.ToString());
        }

        [TestMethod]
        public void EchoWritesTheValuesSeparatedBySpaces()
        {
            var result = Run("WScript.Echo \"a\", 1, True, Null\r\nWScript.Echo\r\nWScript.Echo 1.5");
            Assert.AreEqual("a 1 True Null" + Environment.NewLine + Environment.NewLine + "1.5" + Environment.NewLine, result.Output);
            Assert.AreEqual(0, result.ExitCode);
        }

        [TestMethod]
        public void ArgumentsAndScriptNames()
        {
            var result = Run(
                "WScript.Echo WScript.Arguments.Count & \"|\" & WScript.Arguments(0) & \"|\" & WScript.Arguments.Item(1)\r\n" +
                "Dim a : For Each a In WScript.Arguments : WScript.Echo a : Next\r\n" +
                "WScript.Echo WScript.ScriptName & \"|\" & WScript.ScriptFullName & \"|\" & WScript.Name",
                new[] { "x", "y z" });
            CollectionAssert.AreEqual(new[] { "2|x|y z", "x", "y z", @"test.vbs|C:\scripts\test.vbs|Windows Script Host" }, Lines(result.Output));
        }

        [TestMethod]
        public void QuitEndsTheScriptWithTheExitCodeEvenWithOnErrorResumeNext()
        {
            var result = Run("On Error Resume Next\r\nWScript.Echo \"before\"\r\nWScript.Quit 3\r\nWScript.Echo \"after\"");
            Assert.AreEqual(3, result.ExitCode);
            CollectionAssert.AreEqual(new[] { "before" }, Lines(result.Output));
        }

        [TestMethod]
        public void RuntimeErrorsAreReportedAsCscriptReportsThem()
        {
            var result = Run("WScript.Echo \"start\"\r\nDim x : x = 1 / 0");
            Assert.AreEqual(1, result.ExitCode);
            Assert.AreEqual(@"C:\scripts\test.vbs(2, 1) Microsoft VBScript runtime error: Division by zero", result.Error.Trim());
        }

        [TestMethod]
        public void StandardStreams()
        {
            var result = Run("Dim line : line = WScript.StdIn.ReadLine()\r\nWScript.StdOut.Write \"[\" & line & \"]\"\r\nWScript.StdOut.WriteLine \"!\"\r\nWScript.StdErr.WriteLine \"err\"\r\nWScript.Echo WScript.StdIn.AtEndOfStream", input: "hello\r\n");
            CollectionAssert.AreEqual(new[] { "[hello]!", "True" }, Lines(result.Output));
            Assert.AreEqual("err", result.Error.Trim());
        }

        [TestMethod]
        public void CreateObjectOfWScript()
        {
            var result = Run("Dim d : Set d = WScript.CreateObject(\"Scripting.Dictionary\")\r\nd.Add \"k\", 1\r\nWScript.Echo TypeName(d) & \"|\" & d.Count");
            CollectionAssert.AreEqual(new[] { "Dictionary|1" }, Lines(result.Output));
        }

        [TestMethod]
        public void HelpOfTheCommandLineTool()
        {
            foreach (string option in new[] { "/h", "/help", "/?", "-h" })
            {
                CliRunner.Result result = CliRunner.Run(new[] { option });
                Assert.AreEqual(0, result.ExitCode, option);
                StringAssert.Contains(result.StandardOutput, "[options] scriptname.vbs [script arguments]", option);
                StringAssert.Contains(result.StandardOutput, "/translate", option);
            }
        }

        [TestMethod]
        public void CommandLineToolExitCodesAndErrors()
        {
            string script = Path.Combine(Path.GetTempPath(), "skrypton_cli_" + Guid.NewGuid().ToString("N") + ".vbs");
            File.WriteAllText(script, "WScript.Echo WScript.Arguments(0)\r\nWScript.Quit 4\r\n");
            try
            {
                CliRunner.Result result = CliRunner.Run(new[] { script, "hello world" });
                Assert.AreEqual(4, result.ExitCode);
                Assert.AreEqual("hello world", result.StandardOutput.Trim());

                CliRunner.Result missing = CliRunner.Run(new[] { script + ".missing.vbs" });
                Assert.AreEqual(1, missing.ExitCode);
                StringAssert.Contains(missing.StandardError, "Can not find script file");

                CliRunner.Result unknown = CliRunner.Run(new[] { "/nosuchoption", script });
                Assert.AreEqual(1, unknown.ExitCode);

                CliRunner.Result translated = CliRunner.Run(new[] { "/translate", script });
                Assert.AreEqual(0, translated.ExitCode);
                StringAssert.Contains(translated.StandardOutput, "class Runner");
            }
            finally
            {
                File.Delete(script);
            }
        }

        private static string[] Lines(string text) => CliRunner.Lines(text);
    }
}
