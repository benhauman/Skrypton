using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in WbemScripting.SWbemLocator, also reached through GetObject("winmgmts:"). The classes Win32_Process and Win32_TimeZone are
    /// supported: the processes come from the host's IHostProcessControlHostService (ProcessesCollect) - querying processes is a decision for the
    /// host, so this raises an error if the host has not registered one - and the time zone is the local one of the process (TimeZoneInfo.Local).
    /// Other WMI classes and operations raise a NotSupportedException.
    /// </summary>
    [SourceClassName("SWbemLocator")] // TypeName(CreateObject("WbemScripting.SWbemLocator"))
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonSWbemLocator : ScriptonIReflectOnClrType, IMyMoniker
    {
        public const string ProgId = "WbemScripting.SWbemLocator";

        private readonly IRuntimeHost _runtimeHost;

        public ScriptonSWbemLocator(IRuntimeHost runtimeHost) : base(typeof(ScriptonSWbemLocator))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        /// <summary>
        /// Only the local computer (".", "localhost" or empty) and the root\cimv2 namespace are supported
        /// </summary>
        [DispId(1)]
        public object ConnectServer(
            [Optional, DefaultParameterValue("")] string? strServer,
            [Optional, DefaultParameterValue("")] string? strNamespace,
            [Optional, DefaultParameterValue("")] string? strUser,
            [Optional, DefaultParameterValue("")] string? strPassword,
            [Optional, DefaultParameterValue("")] string? strLocale,
            [Optional, DefaultParameterValue("")] string? strAuthority,
            [Optional, DefaultParameterValue(0)] int iSecurityFlags,
            [Optional, DefaultParameterValue(null)] object? objWbemNamedValueSet)
        {
            string server = (strServer ?? "").Trim();
            if ((server.Length != 0) && (server != ".") && !server.Equals("localhost", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException($"SWbemLocator.ConnectServer: only the local computer is supported, not '{server}'");
            string wmiNamespace = (strNamespace ?? "").Trim();
            if ((wmiNamespace.Length != 0) && !wmiNamespace.Equals(@"root\cimv2", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException($"SWbemLocator.ConnectServer: only the root\\cimv2 namespace is supported, not '{wmiNamespace}'");
            return new ScriptonSWbemServices(_runtimeHost);
        }

        object IMyMoniker.BindToObject(object? pbc, object? pmkToLeft, ref Guid riid)
        {
            return new ScriptonSWbemServices(_runtimeHost);
        }
    }

    [SourceClassName("SWbemServicesEx")] // TypeName(GetObject("winmgmts:")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonSWbemServices : ScriptonIReflectOnClrType
    {
        private static readonly Regex SelectFromWhere = new Regex(
            @"^\s*SELECT\s+.+?\s+FROM\s+(?<class>\w+)(\s+WHERE\s+(?<property>\w+)\s*=\s*'(?<value>[^']*)')?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant
        );

        private readonly IRuntimeHost _runtimeHost;

        public ScriptonSWbemServices(IRuntimeHost runtimeHost) : base(typeof(ScriptonSWbemServices))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        [DispId(0x60020001)]
        public object InstancesOf([In] string strClass, [In, Optional, DefaultParameterValue(0)] int lFlags, [In, Optional, DefaultParameterValue(null)] object? objWbemNamedValueSet)
        {
            if (string.Equals(ScriptonWin32TimeZone.ClassName, strClass, StringComparison.OrdinalIgnoreCase))
                return new ScriptonSWbemObjectSet(new[] { ScriptonWin32TimeZone.FromLocal() });
            return new ScriptonSWbemObjectSet(GetProcesses(strClass));
        }

        /// <summary>
        /// Supports "SELECT ... FROM Win32_Process" with an optional "WHERE Name = '...'" or "WHERE ProcessId = '...'" (compared case-insensitively)
        /// and "SELECT ... FROM Win32_TimeZone"
        /// </summary>
        [DispId(0x60020000)]
        public object ExecQuery([In] string strQuery, [In, Optional, DefaultParameterValue("WQL")] string? strQueryLanguage, [In, Optional, DefaultParameterValue(0)] int lFlags, [In, Optional, DefaultParameterValue(null)] object? objWbemNamedValueSet)
        {
            if (!string.IsNullOrEmpty(strQueryLanguage) && !strQueryLanguage!.Equals("WQL", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException($"SWbemServices.ExecQuery: only WQL is supported, not '{strQueryLanguage}'");

            Match match = SelectFromWhere.Match(strQuery ?? "");
            if (!match.Success)
                throw new NotSupportedException($"SWbemServices.ExecQuery: only \"SELECT ... FROM <class> [WHERE <property> = '<value>']\" is supported, not '{strQuery}'");

            string wmiClass = match.Groups["class"].Value;
            if (string.Equals(ScriptonWin32TimeZone.ClassName, wmiClass, StringComparison.OrdinalIgnoreCase))
            {
                if (match.Groups["property"].Success)
                    throw new NotSupportedException($"SWbemServices.ExecQuery: filtering {ScriptonWin32TimeZone.ClassName} is not supported");
                return InstancesOf(wmiClass);
            }

            IEnumerable<ScriptonWin32Process> instances = GetProcesses(wmiClass);
            if (match.Groups["property"].Success)
            {
                string property = match.Groups["property"].Value;
                string value = match.Groups["value"].Value;
                if (property.Equals("Name", StringComparison.OrdinalIgnoreCase) || property.Equals("Caption", StringComparison.OrdinalIgnoreCase))
                    instances = instances.Where(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase));
                else if (property.Equals("ProcessId", StringComparison.OrdinalIgnoreCase) || property.Equals("Handle", StringComparison.OrdinalIgnoreCase))
                    instances = instances.Where(p => p.ProcessId.ToString(CultureInfo.InvariantCulture) == value);
                else
                    throw new NotSupportedException($"SWbemServices.ExecQuery: filtering Win32_Process by '{property}' is not supported (Name, Caption, ProcessId and Handle are)");
            }
            return new ScriptonSWbemObjectSet(instances.ToList());
        }

        private List<ScriptonWin32Process> GetProcesses(string wmiClass)
        {
            if (!string.Equals("Win32_Process", wmiClass, StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException($"SWbemServices: the WMI class '{wmiClass}' is not supported (only Win32_Process and {ScriptonWin32TimeZone.ClassName} are)");

            IHostProcessControlHostService processControl = _runtimeHost.TryGetRuntimeHostService<IHostProcessControlHostService>()
                ?? throw new InvalidOperationException($"SWbemServices: querying processes requires the host to register an {nameof(IHostProcessControlHostService)}");
            var processes = new List<ScriptonWin32Process>();
            processControl.ProcessesCollect((processId, processName) =>
            {
                processes.Add(new ScriptonWin32Process(processId, processName));
                return true;
            });
            return processes;
        }
    }

    [SourceClassName("SWbemObjectSet")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonSWbemObjectSet : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly IReadOnlyList<object> _items;

        public ScriptonSWbemObjectSet(IEnumerable<object> items) : base(typeof(ScriptonSWbemObjectSet))
        {
            _items = (items ?? throw new ArgumentNullException(nameof(items))).ToList();
        }

        [DispId(1)]
        public int Count => _items.Count;

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _items.GetEnumerator();

        /// <summary>
        /// Item(objectPath) - a Win32_Process.Handle="&lt;id&gt;" path (or a plain process id)
        /// </summary>
        [DispId(0)]
        public object Item(object objIndex)
        {
            string path = Convert.ToString(objIndex, CultureInfo.InvariantCulture) ?? "";
            Match handle = Regex.Match(path, @"(?:Handle|ProcessId)\s*=\s*""?(?<id>\d+)""?", RegexOptions.IgnoreCase);
            string id = handle.Success ? handle.Groups["id"].Value : path.Trim();
            return _items.OfType<ScriptonWin32Process>().FirstOrDefault(p => p.ProcessId.ToString(CultureInfo.InvariantCulture) == id)
                ?? throw new ArgumentException($"SWbemObjectSet.Item: no object with the path '{path}'", nameof(objIndex));
        }
    }

    /// <summary>
    /// A Win32_Process instance (the properties that IHostProcessControlHostService.ProcessesCollect provides)
    /// </summary>
    [SourceClassName("SWbemObjectEx")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWin32Process
    {
        public ScriptonWin32Process(int processId, string name)
        {
            ProcessId = processId;
            Name = name ?? "";
        }

        public string Name { get; }
        public string Caption => Name;
        public int ProcessId { get; }
        public string Handle => ProcessId.ToString(CultureInfo.InvariantCulture);
    }
}
