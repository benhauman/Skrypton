// The members are bound by name from scripts, so they must be instance members with the names of the Windows Script Host (Arguments,
// TextStream, ..)
#pragma warning disable CA1822, CA1716, CA1720, CA1710, CA1711
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.Windows
{
    /// <summary>
    /// The WScript object of the Windows Script Host for scripts run by ScriptonScriptRunner (as with cscript.exe): Echo, Arguments, Quit, Sleep,
    /// StdIn / StdOut / StdErr, CreateObject, GetObject and the script and host names
    /// </summary>
    public sealed class ScriptonWScript
    {
        private readonly Func<string, object> _createObject;
        private readonly Func<string, string, object> _getObject;

        public ScriptonWScript(string scriptFullName, IReadOnlyList<string> arguments, TextWriter standardOutput, TextWriter standardError, TextReader standardInput, Func<string, object> createObject, Func<string, string, object> getObject)
        {
            ScriptFullName = scriptFullName ?? throw new ArgumentNullException(nameof(scriptFullName));
            Arguments = new ScriptonWScriptArguments(arguments ?? throw new ArgumentNullException(nameof(arguments)));
            StdOut = new ScriptonWScriptTextStream(standardOutput ?? throw new ArgumentNullException(nameof(standardOutput)), null);
            StdErr = new ScriptonWScriptTextStream(standardError ?? throw new ArgumentNullException(nameof(standardError)), null);
            StdIn = new ScriptonWScriptTextStream(null, standardInput ?? throw new ArgumentNullException(nameof(standardInput)));
            _createObject = createObject ?? throw new ArgumentNullException(nameof(createObject));
            _getObject = getObject ?? throw new ArgumentNullException(nameof(getObject));
        }

        public string Name => "Windows Script Host";

        public string Version => "5.812";

        public string FullName => Assembly.GetEntryAssembly()?.Location ?? "";

        public string Path => System.IO.Path.GetDirectoryName(FullName) ?? "";

        public string ScriptFullName { get; }

        public string ScriptName => System.IO.Path.GetFileName(ScriptFullName);

        public bool Interactive { get; set; } = true;

        public int Timeout { get; set; }

        public ScriptonWScriptArguments Arguments { get; }

        public ScriptonWScriptTextStream StdOut { get; }

        public ScriptonWScriptTextStream StdErr { get; }

        public ScriptonWScriptTextStream StdIn { get; }

        /// <summary>
        /// Echo [Arg1] [,Arg2] .. - the values separated by spaces, on a line of their own
        /// </summary>
        public void Echo() => StdOut.WriteLine();
        public void Echo(object? a) => EchoValues(a);
        public void Echo(object? a, object? b) => EchoValues(a, b);
        public void Echo(object? a, object? b, object? c) => EchoValues(a, b, c);
        public void Echo(object? a, object? b, object? c, object? d) => EchoValues(a, b, c, d);
        public void Echo(object? a, object? b, object? c, object? d, object? e) => EchoValues(a, b, c, d, e);
        public void Echo(object? a, object? b, object? c, object? d, object? e, object? f) => EchoValues(a, b, c, d, e, f);
        public void Echo(object? a, object? b, object? c, object? d, object? e, object? f, object? g) => EchoValues(a, b, c, d, e, f, g);
        public void Echo(object? a, object? b, object? c, object? d, object? e, object? f, object? g, object? h) => EchoValues(a, b, c, d, e, f, g, h);

        private void EchoValues(params object?[] values) => StdOut.WriteLine(string.Join(" ", values.Select(ToEchoText)));

        /// <summary>
        /// A value as WScript.Echo writes it: as StdOut.Write, but a Boolean is -1 / 0 and Null is "null" (as with Windows Script Host - issue #119)
        /// </summary>
        internal static string ToEchoText(object? value)
        {
            switch (value)
            {
                case bool boolean:
                    return boolean ? "-1" : "0";
                case DBNull _:
                    return "null";
                default:
                    return ToText(value);
            }
        }

        /// <summary>
        /// Quit [ExitCode] - ends the script
        /// </summary>
        public void Quit() => throw new ScriptQuitException(0);

        public void Quit(object? exitCode) => throw new ScriptQuitException((exitCode == null) ? 0 : Convert.ToInt32(exitCode, CultureInfo.InvariantCulture));

        public void Sleep(object milliseconds) => Thread.Sleep(Convert.ToInt32(milliseconds, CultureInfo.InvariantCulture));

        public object CreateObject(string progId) => _createObject(progId);

        public object CreateObject(string progId, string prefix)
        {
            if (!string.IsNullOrEmpty(prefix))
                throw new NotSupportedException($"WScript.CreateObject: connecting the events of '{progId}' to procedures ('{prefix}') is not supported");
            return _createObject(progId);
        }

        public object GetObject(string pathName) => _getObject(pathName, "");

        public object GetObject(string pathName, string progId) => _getObject(pathName, progId);

        internal static string ToText(object? value)
        {
            switch (value)
            {
                case null:
                    return "";
                case DBNull _:
                    return "Null";
                case bool boolean:
                    return boolean ? "True" : "False";
                case DateTime date:
                    return (date.TimeOfDay == TimeSpan.Zero) ? date.ToString("d", CultureInfo.CurrentCulture) : ((date.Date == new DateTime(1899, 12, 30)) ? date.ToString("T", CultureInfo.CurrentCulture) : date.ToString("G", CultureInfo.CurrentCulture));
                // 15 significant digits of a Double and 7 of a Single, as VBScript converts them to text (1 / 3 is 0.333333333333333 - issue #119)
                case double number:
                    return number.ToString("G15", CultureInfo.CurrentCulture);
                case float number:
                    return number.ToString("G7", CultureInfo.CurrentCulture);
                case IFormattable formattable:
                    return formattable.ToString(null, CultureInfo.CurrentCulture);
                default:
                    return value.ToString() ?? "";
            }
        }
    }

    /// <summary>
    /// WScript.Arguments - Item(i) (0-based, the default member), Count, Length
    /// </summary>
    [DefaultMember("Item")]
    [System.Runtime.InteropServices.ComVisible(true)] // the binder follows DefaultMember for COM visible types (WScript.Arguments(0))
    public sealed class ScriptonWScriptArguments : IEnumerable<string>
    {
        private readonly IReadOnlyList<string> _arguments;

        public ScriptonWScriptArguments(IReadOnlyList<string> arguments)
        {
            _arguments = arguments ?? throw new ArgumentNullException(nameof(arguments));
        }

        public string Item(object index)
        {
            int i = Convert.ToInt32(index, CultureInfo.InvariantCulture);
            if ((i < 0) || (i >= _arguments.Count))
                throw new SubscriptOutOfRangeException("WScript.Arguments");
            return _arguments[i];
        }

        public int Count() => _arguments.Count;

        public int Length => _arguments.Count;

        public IEnumerator<string> GetEnumerator() => _arguments.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// WScript.StdOut / StdErr (Write, WriteLine, WriteBlankLines) and WScript.StdIn (Read, ReadLine, ReadAll, SkipLine, AtEndOfStream)
    /// </summary>
    public sealed class ScriptonWScriptTextStream
    {
        private readonly TextWriter? _writer;
        private readonly TextReader? _reader;

        public ScriptonWScriptTextStream(TextWriter? writer, TextReader? reader)
        {
            _writer = writer;
            _reader = reader;
        }

        private TextWriter Writer => _writer ?? throw new InvalidOperationException("Bad file mode (the stream is for reading)");

        private TextReader Reader => _reader ?? throw new InvalidOperationException("Bad file mode (the stream is for writing)");

        public void Write(object? text) => Writer.Write(ScriptonWScript.ToText(text));

        public void WriteLine() => Writer.WriteLine();

        public void WriteLine(object? text) => Writer.WriteLine(ScriptonWScript.ToText(text));

        public void WriteBlankLines(object lines)
        {
            for (int i = Convert.ToInt32(lines, CultureInfo.InvariantCulture); i > 0; i--)
                Writer.WriteLine();
        }

        public string Read(object characters)
        {
            var buffer = new char[Convert.ToInt32(characters, CultureInfo.InvariantCulture)];
            int read = Reader.Read(buffer, 0, buffer.Length);
            return new string(buffer, 0, read);
        }

        public string ReadLine() => Reader.ReadLine() ?? "";

        public string ReadAll() => Reader.ReadToEnd();

        public void SkipLine() => Reader.ReadLine();

        public bool AtEndOfStream => Reader.Peek() < 0;

        public void Close()
        {
            _writer?.Flush();
        }
    }
}
