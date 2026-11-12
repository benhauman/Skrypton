using System;
using System.Globalization;
using System.Linq;
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

        private readonly IRuntimeHost _runtimeHost;
        private readonly IHostEnvironmentHostService _environment;

        public ScriptonWScriptShell(IRuntimeHost runtimeHost) : base(typeof(ScriptonWScriptShell))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
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

        /// <summary>
        /// Environment([Type]) - the variables of "PROCESS", "SYSTEM" (the default), "USER" or "VOLATILE" (issue #107), through the host's
        /// IHostEnvironmentVariablesHostService - or, if the host has registered only its own IHostEnvironmentHostService, read through that (a
        /// sandboxed environment stays sandboxed) - or the process environment
        /// </summary>
        [DispId(200)]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object Environment([In, MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue("SYSTEM")] object? Type)
        {
            string type = IsMissing(Type) ? "SYSTEM" : (Convert.ToString(Type, CultureInfo.InvariantCulture) ?? "").Trim().ToUpperInvariant();
            if ((type != "PROCESS") && (type != "SYSTEM") && (type != "USER") && (type != "VOLATILE"))
                throw new Skrypton.RuntimeSupport.Exceptions.InvalidProcedureCallOrArgumentException($"'WshShell.Environment' (unknown environment '{Type}')");
            IHostEnvironmentVariablesHostService variables = _runtimeHost.TryGetRuntimeHostService<IHostEnvironmentVariablesHostService>()
                ?? ((_runtimeHost.TryGetRuntimeHostService<IHostEnvironmentHostService>() is IHostEnvironmentHostService hostEnvironment)
                    ? new ExpandingEnvironmentVariablesHostService(hostEnvironment)
                    : DefaultHostEnvironmentVariablesHostService.Instance);
            return new ScriptonWshEnvironment(variables, type);
        }

        /// <summary>
        /// LogEvent(Type, Message, [Target]) - an event log entry (issue #116) through the host's IHostEventLogHostService (the Windows Application
        /// log with the source WSH; System.Diagnostics.Trace if the host has none). Type 0 (success), 1 (error), 2 (warning), 4 (information), 8
        /// (audit success) or 16 (audit failure), other types are error 5. Returns True.
        /// </summary>
        [DispId(3000)]
        public bool LogEvent([In, MarshalAs(UnmanagedType.Struct)] object Type, [In, MarshalAs(UnmanagedType.BStr)] string Message, [In, MarshalAs(UnmanagedType.BStr), Optional, DefaultParameterValue("")] string? Target)
        {
            int type = Convert.ToInt32(Type, CultureInfo.InvariantCulture);
            if ((type != 0) && (type != 1) && (type != 2) && (type != 4) && (type != 8) && (type != 16))
                throw new Skrypton.RuntimeSupport.Exceptions.InvalidProcedureCallOrArgumentException($"'WshShell.LogEvent' ({type} is not an event type)");
            IHostEventLogHostService eventLog = _runtimeHost.TryGetRuntimeHostService<IHostEventLogHostService>() ?? DefaultHostEventLogHostService.Instance;
            eventLog.LogEvent(type, Message ?? "", string.IsNullOrEmpty(Target) ? null : Target);
            return true;
        }

        private const string RegReadErrorSource = "WshShell.RegRead";

        /// <summary>
        /// RegRead(Name) - a registry value through the host's IHostRegistryHostService (issue #122), as Windows Script Host returns it: REG_SZ /
        /// REG_EXPAND_SZ a String (not expanded), REG_DWORD a Long, REG_MULTI_SZ an array of Strings and REG_BINARY an array of Bytes (both as
        /// Variant arrays), REG_QWORD is a Type mismatch; a Name that ends with "\" is the default value of the key ("" if it has none). The root
        /// is HKCU, HKLM, HKCR or the full name (also HKEY_USERS, HKEY_CURRENT_CONFIG), case-sensitive; other roots and missing keys and values
        /// raise the errors of WSH.
        /// </summary>
        [DispId(2002)]
        public object RegRead([In, MarshalAs(UnmanagedType.BStr)] string Name)
        {
            string name = Name ?? "";
            int separator = name.IndexOf('\\');
            string? root = (separator < 0) ? null : ToRegistryRoot(name.Substring(0, separator));
            if (root == null)
                throw new Skrypton.RuntimeSupport.Exceptions.CustomException(unchecked((int)0x80070003), RegReadErrorSource, $"Invalid root in registry key \"{name}\".");
            string path = name.Substring(separator + 1);
            int lastSeparator = path.LastIndexOf('\\');
            string subKey = (lastSeparator < 0) ? "" : path.Substring(0, lastSeparator);
            string valueName = (lastSeparator < 0) ? path : path.Substring(lastSeparator + 1);

            IHostRegistryHostService registry = _runtimeHost.TryGetRuntimeHostService<IHostRegistryHostService>() ?? throw new InvalidOperationException(
                $"WshShell.RegRead: reading the registry requires the host to register an {nameof(IHostRegistryHostService)}");
            HostRegistryValue? value = registry.ReadValue(root, subKey, valueName, out bool keyExists);
            if (!keyExists)
                throw new Skrypton.RuntimeSupport.Exceptions.CustomException(unchecked((int)0x80070002), RegReadErrorSource, $"Invalid root in registry key \"{name}\".");
            if (value == null)
            {
                if (valueName.Length == 0)
                    return ""; // the default value of a key that has none
                throw new Skrypton.RuntimeSupport.Exceptions.CustomException(unchecked((int)0x80070002), RegReadErrorSource, $"Unable to open registry key \"{name}\" for reading.");
            }
            switch (value.Kind)
            {
                case HostRegistryValueKind.DWord:
                    return Convert.ToInt32(value.Data, CultureInfo.InvariantCulture);
                case HostRegistryValueKind.MultiSz:
                    return ((string[])value.Data).Cast<object>().ToArray();
                case HostRegistryValueKind.Binary:
                    return ((byte[])value.Data).Cast<object>().ToArray();
                case HostRegistryValueKind.Sz:
                case HostRegistryValueKind.ExpandSz:
                    return Convert.ToString(value.Data, CultureInfo.InvariantCulture) ?? "";
                default:
                    throw new Skrypton.RuntimeSupport.TypeMismatchException($"'WshShell.RegRead' ({name} is a {value.Kind} value)");
            }
        }

        private static string? ToRegistryRoot(string root)
        {
            // case-sensitive, as with Windows Script Host ("hklm" is an invalid root); HKU and HKCC are not abbreviations WSH knows
            switch (root)
            {
                case "HKCU":
                case "HKEY_CURRENT_USER":
                    return "HKEY_CURRENT_USER";
                case "HKLM":
                case "HKEY_LOCAL_MACHINE":
                    return "HKEY_LOCAL_MACHINE";
                case "HKCR":
                case "HKEY_CLASSES_ROOT":
                    return "HKEY_CLASSES_ROOT";
                case "HKEY_USERS":
                case "HKEY_CURRENT_CONFIG":
                    return root;
                default:
                    return null;
            }
        }

        private IHostProcessControlHostService ProcessControl(string member)
        {
            return _runtimeHost.TryGetRuntimeHostService<IHostProcessControlHostService>() ?? throw new InvalidOperationException(
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
