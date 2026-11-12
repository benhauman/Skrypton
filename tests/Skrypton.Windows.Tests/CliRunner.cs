#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// Runs the command line tool of the test project (Skrypton.Windows.Cli or Skrypton.Windows.NetFx.Cli, see AssemblyMetadata "CliPath")
    /// </summary>
    internal static class CliRunner
    {
        public static string CliPath => typeof(CliRunner).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>().Single(a => a.Key == "CliPath").Value!;

        public static string RepositoryRoot => typeof(CliRunner).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>().Single(a => a.Key == "RepositoryRoot").Value!;

        public sealed class Result
        {
            public Result(int exitCode, string standardOutput, string standardError)
            {
                ExitCode = exitCode;
                StandardOutput = standardOutput;
                StandardError = standardError;
            }

            public int ExitCode { get; }
            public string StandardOutput { get; }
            public string StandardError { get; }
        }

        /// <summary>
        /// Run the command line tool in the test process (Program.Run, its code with the same options): a process for each script would
        /// compile Roslyn with the JIT each time - about 1 s (.NET) or 2 s (.NET Framework) of CPU, 10 s and more on the CI agent (issue #136).
        /// The environment variables are set in the process (the tests use their own names or the same values).
        /// </summary>
        public static Result Run(IEnumerable<string> arguments, IDictionary<string, string>? environment = null)
        {
            if (environment != null)
            {
                foreach (KeyValuePair<string, string> variable in environment)
                    Environment.SetEnvironmentVariable(variable.Key, variable.Value);
            }
            var standardOutput = new StringWriter();
            var standardError = new StringWriter();
            System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.CurrentCulture; // /culture sets it
            try
            {
                int exitCode = Skrypton.Windows.Cli.Program.Run(Path.GetFileNameWithoutExtension(CliPath), arguments.ToArray(), standardOutput, standardError, TextReader.Null);
                return new Result(exitCode, standardOutput.ToString(), standardError.ToString());
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = culture;
            }
        }

        /// <summary>
        /// Start the command line tool (the tests of the tool itself: options, exit codes, the environment of its process)
        /// </summary>
        public static Result RunProcess(IEnumerable<string> arguments)
        {
            if (!File.Exists(CliPath))
                throw new InvalidOperationException("The command line tool has not been built: " + CliPath);
            var startInfo = new ProcessStartInfo(CliPath, string.Join(" ", arguments.Select(Quote)))
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            using (Process process = Process.Start(startInfo)!)
            {
                process.StandardInput.Close();
                var standardError = process.StandardError.ReadToEndAsync();
                string standardOutput = process.StandardOutput.ReadToEnd();
                if (!process.WaitForExit(5 * 60 * 1000))
                {
                    process.Kill();
                    throw new TimeoutException("The command line tool did not end within 5 minutes");
                }
                return new Result(process.ExitCode, standardOutput, standardError.Result);
            }
        }

        private static string Quote(string argument) => ((argument.Length == 0) || argument.Any(char.IsWhiteSpace) || argument.Contains("\"")) ? "\"" + argument.Replace("\"", "\\\"") + "\"" : argument;

        public static string[] Lines(string text) => text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
    }
}
