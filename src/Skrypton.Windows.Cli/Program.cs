using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Skrypton.Windows;

namespace Skrypton.Windows.Cli
{
    /// <summary>
    /// Runs VBScript files from the command line like cscript.exe (shared by Skrypton.Windows.Cli and Skrypton.Windows.NetFx.Cli)
    /// </summary>
    internal static class Program
    {
        private static readonly char[] ProgIdSeparators = { ',', ';' };

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8; // before Console.Out is passed on: setting the encoding replaces Console.Out
            string exe = Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly()?.Location ?? "Skrypton.Windows.Cli");
            return Run(exe, args, Console.Out, Console.Error, Console.In);
        }

        /// <summary>
        /// The command line tool with these streams (the tests run the parity scripts in their process with it - issue #136)
        /// </summary>
        internal static int Run(string exe, string[] args, TextWriter output, TextWriter error, TextReader input)
        {
            var options = new List<string>();
            int index = 0;
            while ((index < args.Length) && (args[index].StartsWith("/", StringComparison.Ordinal) || args[index].StartsWith("-", StringComparison.Ordinal)))
                options.Add(args[index++].TrimStart('/', '-').ToLowerInvariant());

            if (options.Any(o => (o == "h") || (o == "help") || (o == "?")) || (index >= args.Length))
            {
                output.Write(Usage(exe));
                return (index >= args.Length) && (options.Count == 0) ? 1 : 0;
            }

            string? unknown = options.FirstOrDefault(o => (o != "nologo") && (o != "logo") && (o != "translate") && !o.StartsWith("culture:", StringComparison.Ordinal) && !o.StartsWith("com:", StringComparison.Ordinal));
            if (unknown != null)
            {
                error.WriteLine($"Input Error: Unknown option \"/{unknown}\" specified (see {exe} /help).");
                return 1;
            }

            CultureInfo culture = CultureInfo.CurrentCulture;
            string? cultureOption = options.LastOrDefault(o => o.StartsWith("culture:", StringComparison.Ordinal));
            if (cultureOption != null)
            {
                culture = CultureInfo.GetCultureInfo(cultureOption.Substring("culture:".Length));
                CultureInfo.CurrentCulture = culture;
            }

            if (options.Contains("logo"))
                output.WriteLine(Logo());

            // /com:<ProgID>[,<ProgID>..] or /com:* - CreateObject also creates these registered COM classes (issue #103)
            WindowsRuntimeHost host = WindowsRuntimeHost.CreateDefault();
            string[] comProgIds = options.Where(o => o.StartsWith("com:", StringComparison.Ordinal))
                .SelectMany(o => o.Substring("com:".Length).Split(ProgIdSeparators, StringSplitOptions.RemoveEmptyEntries))
                .Select(p => p.Trim())
                .ToArray();
            if (comProgIds.Contains("*"))
                host.Register<Skrypton.RuntimeSupport.Implementations.IHostComActivationHostService>(new WindowsHostComActivationHostService(_ => true));
            else if (comProgIds.Length > 0)
                host.RegisterComActivation(comProgIds);

            var runner = new ScriptonScriptRunner(host, output, error, input, culture);
            string script = args[index];
            if (options.Contains("translate"))
            {
                if (!File.Exists(script))
                {
                    error.WriteLine($"Input Error: Can not find script file \"{Path.GetFullPath(script)}\".");
                    return 1;
                }
                output.Write(runner.Translate(ScriptonScriptRunner.ReadScriptFile(script)));
                return 0;
            }
            return runner.RunFile(script, args.Skip(index + 1).ToList());
        }

        private static string Logo()
        {
            Assembly assembly = typeof(ScriptonScriptRunner).Assembly;
            string version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? assembly.GetName().Version?.ToString() ?? "";
            return $"Skrypton Script Host version {version} ({RuntimeDescription()})" + Environment.NewLine + "VBScript translated to C#, compiled with Roslyn and run on .NET." + Environment.NewLine;
        }

        private static string RuntimeDescription() => System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;

        private static string Usage(string exe)
        {
            return Logo() + Environment.NewLine +
                $"Usage: {exe} [options] scriptname.vbs [script arguments]" + Environment.NewLine +
                Environment.NewLine +
                "Runs a VBScript file like cscript.exe: the script is translated to C#, compiled and run. WScript.Echo writes to the console," + Environment.NewLine +
                "WScript.Arguments are the script arguments and WScript.Quit sets the exit code. CreateObject provides Scripting.FileSystemObject," + Environment.NewLine +
                "Scripting.Dictionary, VBScript.RegExp, WScript.Shell, Shell.Application, ADODB, MSXML, WMI (winmgmts:), Word, Excel, Outlook," + Environment.NewLine +
                "InternetExplorer.Application and htmlfile with the services of Windows (processes, clipboard, dialogs, OLE DB, Outlook)." + Environment.NewLine +
                Environment.NewLine +
                "Options (before the script name, '/' or '-'):" + Environment.NewLine +
                "  /h, /help, /?       Show this help" + Environment.NewLine +
                "  /logo               Show the banner before running the script" + Environment.NewLine +
                "  /nologo             Do not show the banner (the default, accepted for cscript compatibility)" + Environment.NewLine +
                "  /translate          Write the C# program that the script is translated to instead of running it" + Environment.NewLine +
                "  /culture:<name>     Run with the culture (eg. /culture:de-DE) instead of the culture of the user" + Environment.NewLine +
                "  /com:<ProgIDs>      CreateObject also creates the registered COM classes of these ProgIDs (comma separated, * for all) -" + Environment.NewLine +
                "                      built-in objects come first (eg. /com:WScript.Network,hlStartITIL.Global)" + Environment.NewLine +
                Environment.NewLine +
                "Exit code: the code of WScript.Quit, 0 when the script ends, 1 after an error (reported as cscript reports it, eg." + Environment.NewLine +
                "\"C:\\scripts\\x.vbs(3, 1) Microsoft VBScript runtime error: Type mismatch\")." + Environment.NewLine +
                Environment.NewLine +
                "Examples:" + Environment.NewLine +
                $"  {exe} report.vbs" + Environment.NewLine +
                $"  {exe} /culture:de-DE import.vbs C:\\data\\input.csv" + Environment.NewLine +
                $"  {exe} /translate report.vbs > report.cs" + Environment.NewLine;
        }
    }
}
