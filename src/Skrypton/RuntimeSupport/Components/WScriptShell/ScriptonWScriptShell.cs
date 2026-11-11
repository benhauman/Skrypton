using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in WScript.Shell (WshShell). Processes are started through the host's IHostProcessControlHostService - starting processes from a
    /// script is a decision for the host, so Run / Exec / AppActivate raise an error if the host has not registered one. Environment variables are
    /// resolved through the host's IHostEnvironmentHostService, or the process environment if there is none.
    /// </summary>
    [SourceClassName("IWshShell3")] // TypeName(CreateObject("WScript.Shell")) is "IWshShell3" (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWScriptShell : ScriptonIReflectOnClrType
    {
        public const string ProgId = "WScript.Shell";

        private readonly IHostProcessControlHostService? _processControl;
        private readonly IHostEnvironmentHostService _environment;

        public ScriptonWScriptShell(IRuntimeHost runtimeHost) : base(typeof(ScriptonWScriptShell))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _processControl = runtimeHost.TryGetRuntimeHostService<IHostProcessControlHostService>();
            _environment = runtimeHost.TryGetRuntimeHostService<IHostEnvironmentHostService>() ?? DefaultHostEnvironmentHostService.Instance;
        }

        /// <summary>
        /// Run(Command, [WindowStyle = 1], [WaitOnReturn = False]) - returns the exit code if WaitOnReturn is True, otherwise 0
        /// </summary>
        [DispId(1000)]
        public int Run([In, MarshalAs(UnmanagedType.BStr)] string Command, [In, MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(1)] object? WindowStyle, [In, MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(false)] object? WaitOnReturn)
        {
            if (string.IsNullOrWhiteSpace(Command))
                throw new ArgumentException("WshShell.Run: a command is required", nameof(Command));

            // 0 hidden, 1 normal (default), 2 minimized, 3 maximized, 4 normal without activating, 5 show, 6 minimize, 7 minimized without activating, ..
            byte windowStyle = IsMissing(WindowStyle) ? (byte)1 : Convert.ToByte(WindowStyle, CultureInfo.InvariantCulture);
            bool waitOnReturn = !IsMissing(WaitOnReturn) && ToBool(WaitOnReturn!);
            return ProcessControl("Run").ProcessStart(Command, windowStyle, waitOnReturn);
        }

        /// <summary>
        /// Exec(Command) - runs the command line application and returns a WshScriptExec with its exit code and output (the application has finished
        /// when this returns, so Status is 1)
        /// </summary>
        [DispId(1009)]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object Exec([In, MarshalAs(UnmanagedType.BStr)] string Command)
        {
            if (string.IsNullOrWhiteSpace(Command))
                throw new ArgumentException("WshShell.Exec: a command is required", nameof(Command));
            return new ScriptonWshScriptExec(ProcessControl("Exec").ProcessExec(Command));
        }

        /// <summary>
        /// AppActivate(App) - App is a process id (activating by window title is not supported)
        /// </summary>
        [DispId(1012)]
        public bool AppActivate([MarshalAs(UnmanagedType.Struct)] object App, [MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(null)] object? Wait)
        {
            if ((App is string title) && !int.TryParse(title, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                throw new NotSupportedException($"WshShell.AppActivate: activating a window by its title ('{title}') is not supported, pass a process id");
            return ProcessControl("AppActivate").ProcessActivate(Convert.ToInt32(App, CultureInfo.InvariantCulture));
        }

        [DispId(1006)]
        public string ExpandEnvironmentStrings([In, MarshalAs(UnmanagedType.BStr)] string Src)
        {
            return _environment.ExpandEnvironmentStrings(Src ?? "");
        }

        private IHostProcessControlHostService ProcessControl(string member)
        {
            return _processControl ?? throw new InvalidOperationException(
                $"WshShell.{member}: starting or controlling processes requires the host to register an {nameof(IHostProcessControlHostService)}"
            );
        }

        private static bool IsMissing(object? value) => (value == null) || (value is System.Reflection.Missing);

        private static bool ToBool(object value)
        {
            if (value is bool boolean)
                return boolean;
            if (value is string text)
                return bool.TryParse(text, out bool parsed) ? parsed : (Convert.ToDouble(text, CultureInfo.InvariantCulture) != 0);
            return Convert.ToDouble(value, CultureInfo.InvariantCulture) != 0;
        }
    }

    /// <summary>
    /// The WshScriptExec returned by WshShell.Exec - the application has already finished (see ScriptonWScriptShell.Exec)
    /// </summary>
    [SourceClassName("WshScriptExec")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWshScriptExec : ScriptonIReflectOnClrType
    {
        private readonly HostProcessExecResult _result;

        public ScriptonWshScriptExec(HostProcessExecResult result) : base(typeof(ScriptonWshScriptExec))
        {
            _result = result ?? throw new ArgumentNullException(nameof(result));
            StdOut = new ScriptonTextStream(result.StdOut);
            StdErr = new ScriptonTextStream(result.StdErr);
        }

        /// <summary>
        /// 1 (WshFinished)
        /// </summary>
        [DispId(1)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public int Status => 1;
#pragma warning restore CA1822

        [DispId(2)]
        public int ExitCode => _result.ExitCode;

        [DispId(3)]
        public int ProcessID => _result.ProcessId;

        [DispId(4)]
        public object StdOut { get; }

        [DispId(5)]
        public object StdErr { get; }

        /// <summary>
        /// The application has already finished, so there is nothing to terminate
        /// </summary>
        [DispId(6)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public void Terminate()
#pragma warning restore CA1822
        {
        }
    }

    /// <summary>
    /// A read-only TextStream over captured text (WshScriptExec.StdOut / StdErr)
    /// </summary>
    [SourceClassName("TextStream")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonTextStream : ScriptonIReflectOnClrType
    {
        private readonly string _text;
        private int _position;

        public ScriptonTextStream(string text) : base(typeof(ScriptonTextStream))
        {
            _text = text ?? "";
        }

        [DispId(1)]
        public bool AtEndOfStream => _position >= _text.Length;

        [DispId(2)]
        public string ReadAll()
        {
            string rest = _text.Substring(_position);
            _position = _text.Length;
            return rest;
        }

        [DispId(3)]
        public string ReadLine()
        {
            if (AtEndOfStream)
                return "";
            int end = _text.IndexOf('\n', _position);
            string line;
            if (end < 0)
            {
                line = _text.Substring(_position);
                _position = _text.Length;
            }
            else
            {
                line = _text.Substring(_position, end - _position);
                _position = end + 1;
            }
            return line.TrimEnd('\r');
        }

        [DispId(4)]
        public string Read(int Characters)
        {
            int count = Math.Max(0, Math.Min(Characters, _text.Length - _position));
            string read = _text.Substring(_position, count);
            _position += count;
            return read;
        }
    }
}
