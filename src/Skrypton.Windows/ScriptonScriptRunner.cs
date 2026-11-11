using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Windows
{
    /// <summary>
    /// Runs VBScript files as cscript.exe does: the script is translated to C#, compiled and run with a WScript object (Echo writes to the
    /// standard output) and the host services of the runtime host (WindowsRuntimeHost.CreateDefault() for the services of Windows). Runtime
    /// errors are reported as cscript reports them ("script.vbs(3, 1) Microsoft VBScript runtime error: Type mismatch").
    /// </summary>
    public sealed class ScriptonScriptRunner
    {
        private readonly IRuntimeHost _runtimeHost;
        private readonly TextWriter _standardOutput;
        private readonly TextWriter _standardError;
        private readonly TextReader _standardInput;
        private readonly CultureInfo _culture;

        public ScriptonScriptRunner(IRuntimeHost runtimeHost, TextWriter standardOutput, TextWriter standardError, TextReader standardInput, CultureInfo culture)
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
            _standardOutput = standardOutput ?? throw new ArgumentNullException(nameof(standardOutput));
            _standardError = standardError ?? throw new ArgumentNullException(nameof(standardError));
            _standardInput = standardInput ?? throw new ArgumentNullException(nameof(standardInput));
            _culture = culture ?? throw new ArgumentNullException(nameof(culture));
        }

        /// <summary>
        /// Run the script file - the exit code is the one of WScript.Quit, 0 when the script ends, 1 after an error
        /// </summary>
        public int RunFile(string scriptPath, IReadOnlyList<string> arguments)
        {
            string fullPath = Path.GetFullPath(scriptPath ?? throw new ArgumentNullException(nameof(scriptPath)));
            if (!File.Exists(fullPath))
            {
                _standardError.WriteLine($"Input Error: Can not find script file \"{fullPath}\".");
                return 1;
            }
            return Run(ReadScriptFile(fullPath), fullPath, arguments);
        }

        /// <summary>
        /// Run the script source (scriptFullName is WScript.ScriptFullName and is used in error messages)
        /// </summary>
        public int Run(string source, string scriptFullName, IReadOnlyList<string> arguments)
        {
            var logger = new RuntimeLogger();
            IScriptControl control = CreateScriptControl(logger);
            var wscript = new ScriptonWScript(scriptFullName, arguments ?? Array.Empty<string>(), _standardOutput, _standardError, _standardInput, CreateObject, GetObject);
            control.AddObject("WScript", wscript, false);
            try
            {
                control.ExecuteStatement(NormalizeLineEnds(source ?? ""));
                return 0;
            }
            catch (Exception e) when (ScriptQuitException.Find(e) is ScriptQuitException quit)
            {
                return quit.ExitCode;
            }
            catch (Exception e)
            {
                Error? error = control.Error;
                if ((error != null) && (error.Number != 0))
                    _standardError.WriteLine($"{scriptFullName}({error.Line}, {error.Column + 1}) {error.Source}: {error.Description}");
                else
                    _standardError.WriteLine($"{scriptFullName}(0, 1) {e.GetType().Name}: {e.Message}");
                return 1;
            }
            finally
            {
                _standardOutput.Flush();
            }
        }

        /// <summary>
        /// The C# program that the script is translated to
        /// </summary>
        public string Translate(string source)
        {
            return Skrypton.CSharpWriter.DefaultTranslator.TranslateExecutable(
                _culture,
                NormalizeLineEnds(source ?? ""),
                HostObjectNames,
                Array.Empty<Skrypton.CSharpWriter.CodeTranslation.BlockTranslators.ExternalMemberMethodInfo>(),
                Array.Empty<string>(),
                undeclaredVariablesAreErrors: false);
        }

        private static readonly string[] HostObjectNames = { "WScript" };

        private IScriptControl CreateScriptControl(IRuntimeLogger logger)
        {
            // Scripts written for cscript compile with warnings (unused variables, code after Exit Sub, ..), so warnings are not errors
            var configuration = new ScriptControlConfiguration(tempEnabled: false, tempDirectoryPath: null, enabledLoadFromDisk: false, translationSuppression: Array.Empty<string>(), noWarn: Array.Empty<string>(), treatWarningsAsErrors: false, undeclaredVariablesAreErrors: false);
            IScriptControl control = new ScriptControlClass(_runtimeHost, logger, _culture, configuration);
            control.Language = "VBScript";
            control.AllowUI = true;
            control.Timeout = -1;
            return control;
        }

        private object CreateObject(string progId) => CompatLayer().CREATEOBJECT(progId);

        private object GetObject(string pathName, string progId) => string.IsNullOrEmpty(progId) ? CompatLayer().GETOBJECT(pathName) : CompatLayer().GETOBJECT(pathName, progId);

        private DefaultRuntimeFunctionalityProvider CompatLayer()
        {
            return (DefaultRuntimeFunctionalityProvider)DefaultRuntimeSupportClassFactory.Create(_runtimeHost, new RuntimeLogger(), _culture).Get();
        }

        /// <summary>
        /// Read the script as cscript does: Unicode with a byte order mark, otherwise the ANSI code page of the culture
        /// </summary>
        public static string ReadScriptFile(string path)
        {
            byte[] content = File.ReadAllBytes(path);
            if ((content.Length >= 3) && (content[0] == 0xEF) && (content[1] == 0xBB) && (content[2] == 0xBF))
                return new UTF8Encoding(false).GetString(content, 3, content.Length - 3);
            if ((content.Length >= 2) && (content[0] == 0xFF) && (content[1] == 0xFE))
                return Encoding.Unicode.GetString(content, 2, content.Length - 2);
            if ((content.Length >= 2) && (content[0] == 0xFE) && (content[1] == 0xFF))
                return Encoding.BigEndianUnicode.GetString(content, 2, content.Length - 2);
            return AnsiEncoding().GetString(content);
        }

        private static Encoding AnsiEncoding()
        {
#if NETFRAMEWORK
            return Encoding.Default;
#else
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
#endif
        }

        private static string NormalizeLineEnds(string source) => source.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");

        private sealed class RuntimeLogger : IRuntimeLogger
        {
            public void LogException(Exception exception)
            {
            }
        }
    }
}
