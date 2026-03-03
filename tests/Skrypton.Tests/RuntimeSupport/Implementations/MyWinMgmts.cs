using Skrypton.RuntimeSupport.Attributes;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Implementations;
using System.ComponentModel;
using System.Collections;
using System.Linq;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    // wbemdisp.tlb

    [SourceClassName("SWbemLocator")] // for TYPENAME(CreateObject("WbemScripting.SWbemLocator"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MySWbemLocator : IReflectOnClrType, ISWbemLocator, IMyMoniker
    {
        /*
Set locator = CreateObject("WbemScripting.SWbemLocator")
Set objWMI = locator.ConnectServer(".", "root\cimv2")
WScript.Echo TypeName(objWMI)  ' Outputs: SWbemServices
         */
        private readonly IServiceProvider _hostServices;
        public MySWbemLocator(IServiceProvider hostServices, string optionalDisplayName) : base(typeof(MySWbemLocator)) // see IClassFactory2::CreateInstanceMoniker
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        public ISWbemServices ConnectServer(string strServer, string strNamespace, string strUser, string strPassword, string strLocale, string strAuthority, int iSecurityFlags, object objWbemNamedValueSet)
        {
            return new MySWbemServices(_hostServices);
        }

        object IMyMoniker.BindToObject(object pbc, object pmkToLeft, ref Guid riid)
        {
            return ConnectServer(strServer: null, strNamespace: null, strUser: null, strPassword: null, strLocale: null, strAuthority: null, iSecurityFlags: 0, objWbemNamedValueSet: null);
        }
    }

    [SourceClassName("SWbemServices")] // for TYPENAME(GetObject("winmgmts:"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[DefaultMember("Item")]
    internal sealed class MySWbemServices : IReflectOnClrType, ISWbemServices
    {
        /*
         1. Set obj = GetObject("winmgmts:")
            MkParseDisplayName (ole32.dll) => Moniker! "winmgmts:" => COM uses 'winmgmts' as a progid => HKEY_CLASSES_ROOT\CLSID\{172BDDF8-CEEA-11D1-8B05-00600806D9B6}\VersionIndependentProgID => WbemScripting.SWbemLocator in wbemdisp.dll
            IMoniker::BindToObject


        c#:
        net48: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Management.dll
        net10: dotnet add package System.Management =>
        static void Main()
        {
            // Connect to the local WMI namespace
            ConnectionOptions options = new ConnectionOptions
            {
                Username = "user",
                Password = "password"
            };
            ManagementScope remoteScope = new ManagementScope(@"\\remoteMachine\root\cimv2", options);
            remoteScope.Connect();

        OR
            ManagementScope scope = new ManagementScope(@"\\.\root\cimv2");
            scope.Connect();

            // Query something, e.g., OS information
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            foreach (ManagementObject os in searcher.Get())
            {
                Console.WriteLine($"OS Name: {os["Caption"]}");
                Console.WriteLine($"Version: {os["Version"]}");
            }
        }
         */
        private readonly IServiceProvider _hostServices;
        public MySWbemServices(IServiceProvider hostServices) : base(typeof(MySWbemServices))
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        public ISWbemSecurity Security_ => throw new NotImplementedException();

        public ISWbemObject ExecMethod([In] string strObjectPath, [In] string strMethodName, [In, Optional] object objInParameters, [In, Optional] int lFlags, [In, Optional] object objWbemNamedValueSet)
        {
            throw new NotImplementedException();
        }

        public ISWbemObjectSet ExecNotificationQuery([In] string strQuery, [In, Optional] string strQueryLanguage, [In, Optional] int lFlags, [In, Optional] object objWbemNamedValueSet)
        {
            throw new NotImplementedException();
        }

        [DispId(1610743809)]
        public ISWbemObjectSet ExecQuery(string strQuery, string strQueryLanguage = "WQL", int lFlags = 0, object objWbemNamedValueSet = null)
        {
            throw new NotImplementedException();
        }

        public ISWbemObject Get([In] string strObjectPath, [In, Optional] int lFlags, [In, Optional] object objWbemNamedValueSet)
        {
            throw new NotImplementedException();
        }

        public ISWbemObject GetObjectText([In] string strObjectPath, [In, Optional] int iFlags)
        {
            throw new NotImplementedException();
        }

        public ISWbemObjectSet InstancesOf([In] string strClass, [In, Optional, DefaultParameterValue(0)] int lFlags, [In, Optional, DefaultParameterValue(null)] object objWbemNamedValueSet)
        {
            if (string.Equals("Win32_Process", strClass, StringComparison.OrdinalIgnoreCase))
            {
                var svc = _hostServices.GetRequiredService<IHostProcessControlHostService>();

                List<ProcessInfo> processes = new List<ProcessInfo>();
                svc.ProcessesCollect((int processId, string processName) =>
                {
                    processes.Add(new ProcessInfo(processId, processName));
                    return true; // continue
                });
                //return new MySWbemObjectSet(new ProcessInfo[] { new ProcessInfo(666, "xxx"), new ProcessInfo(4444, "IExplore.exe") }.Select(x => (object)x).ToArray());
                //return new MySWbemObjectSet(Array.Empty<object>());
                return new MySWbemObjectSet(processes.Select(x => (object)x).ToArray());
            }
            throw new NotImplementedException();
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ProcessInfo
    {
        public ProcessInfo(int processId, string name)
        {
            ProcessId = processId;
            Name = name;
        }

        //[DispId(0)]
        public string Name { get; }

        public int ProcessId { get; }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MySWbemObjectSet : IReflectOnClrType, ISWbemObjectSet
    {
        private readonly IList<object> _items;
        public MySWbemObjectSet(IList<object> items) : base(typeof(MySWbemObjectSet))
        {
            _items = items;
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        [DispId(0)]
        public ISWbemObject Item(object objIndex)
        {
            throw new NotImplementedException();
        }
    }

    //[ComImport, Guid("76A64158-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemLocator
    {
        [DispId(1610743809)]
        ISWbemServices ConnectServer(
            string strServer,
            string strNamespace,
            string strUser,
            string strPassword,
            string strLocale,
            string strAuthority,
            int iSecurityFlags,
            object objWbemNamedValueSet);
    }

    //[ComImport, Guid("76A6415A-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemServices
    {
        [DispId(0x60020000)]
        ISWbemObjectSet ExecQuery(
            [In] string strQuery,
            [In, Optional, DefaultValue("WQL")] string strQueryLanguage,// = "WQL",
            [In, Optional] int lFlags,// = 0,
            [In, Optional] object objWbemNamedValueSet// = null
        );

        // Retrieves instances of a class
        [DispId(0x60020001)]
        ISWbemObjectSet InstancesOf(
            [In] string strClass,
            [In, Optional, DefaultParameterValue(0)] int lFlags,// = 0, , DefaultValue(0)
            [In, Optional, DefaultParameterValue(null)] object objWbemNamedValueSet// = null
        );

        // Retrieves a single object by path
        [DispId(0x60020002)]
        ISWbemObject Get(
            [In] string strObjectPath,
            [In, Optional] int lFlags,// = 0,
            [In, Optional] object objWbemNamedValueSet// = null
        );

        // Executes a WMI method on a given object
        [DispId(0x60020003)]
        ISWbemObject ExecMethod(
            [In] string strObjectPath,
            [In] string strMethodName,
            [In, Optional] object objInParameters,// = null,
            [In, Optional] int lFlags,// = 0,
            [In, Optional] object objWbemNamedValueSet// = null
        );

        // Other methods include:
        [DispId(0x60020004)]
        ISWbemObject GetObjectText(
            [In] string strObjectPath,
            [In, Optional] int iFlags// = 0
        );

        [DispId(0x60020005)]
        ISWbemObjectSet ExecNotificationQuery(
            [In] string strQuery,
            [In, Optional] string strQueryLanguage,// = "WQL",
            [In, Optional] int lFlags,// = 0,
            [In, Optional] object objWbemNamedValueSet// = null
        );

        // Security settings
        [DispId(0x60020006)]
        ISWbemSecurity Security_ { get; }
    }

    //[ComImport]
    //[Guid("76A64156-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemSecurity
    {
        // Gets/sets impersonation level
        [DispId(1)]
        int ImpersonationLevel { get; set; }

        // Gets/sets authentication level
        [DispId(2)]
        int AuthenticationLevel { get; set; }

        // Gets/sets privileges
        [DispId(3)]
        int Privileges { get; set; }

        // Gets/sets security descriptor
        [DispId(4)]
        object SecurityDescriptor { get; set; }

        // Gets/sets whether DCOM authentication is required
        [DispId(5)]
        bool EnablePrivileges { get; set; }
    }

    //[ComImport, Guid("76A64162-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemObjectSet
    {
        [DispId(0)]
        ISWbemObject Item(object objIndex);

        [DispId(-4)]
        System.Collections.IEnumerator GetEnumerator();

        // Other members: Count, _NewEnum, etc.
    }

    //[ComImport]
    //[Guid("76A64163-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemObject
    {
        // Properties
        [DispId(0x00000001)]
        ISWbemProperties Properties_ { get; }

        [DispId(0x00000002)]
        ISWbemMethods Methods_ { get; }

        [DispId(0x00000003)]
        ISWbemQualifierSet Qualifiers_ { get; }

        [DispId(0x00000004)]
        string Path_ { get; }

        [DispId(0x00000005)]
        string Security_ { get; }

        // Methods
        [DispId(0x00000006)]
        string GetObjectText_(int iFlags);

        [DispId(0x00000007)]
        ISWbemObject Clone_();

        [DispId(0x00000008)]
        ISWbemObject Put_(int iFlags);

        [DispId(0x00000009)]
        void Delete_(int iFlags);

        [DispId(0x0000000A)]
        ISWbemObject ExecMethod_(string strMethodName, object objInParams, int iFlags);
    }

    //[ComImport]
    //[Guid("76A64164-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemProperties
    {
        [DispId(0)]
        ISWbemProperty Item(object strName);  // Access property by name

        [DispId(-4)]
        System.Collections.IEnumerator GetEnumerator();

        [DispId(1)]
        int Count { get; }
    }

    //[ComImport]
    //[Guid("76A64165-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemProperty
    {
        [DispId(0)]
        string Name { get; }

        [DispId(1)]
        object Value { get; set; }

        [DispId(2)]
        string CIMType { get; }
    }

    //[ComImport]
    //[Guid("76A64166-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemMethods
    {
        [DispId(0)]
        ISWbemMethod Item(object strName);  // Access method by name

        [DispId(-4)]
        System.Collections.IEnumerator GetEnumerator();

        [DispId(1)]
        int Count { get; }
    }

    //[ComImport]
    //[Guid("76A64167-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemMethod
    {
        [DispId(0)]
        string Name { get; }

        [DispId(1)]
        string InParameters { get; }   // Returns WMI object representing input parameters

        [DispId(2)]
        string OutParameters { get; }  // Returns WMI object representing output parameters
    }

    //[ComImport]
    //[Guid("76A64168-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemQualifierSet
    {
        [DispId(0)]
        ISWbemQualifier Item(string strName);

        [DispId(-4)]
        System.Collections.IEnumerator GetEnumerator();

        [DispId(1)]
        int Count { get; }

        [DispId(2)]
        void Add(string strName, object varVal, bool bPropagatesToSubclass, bool bPropagatesToInstance);

        [DispId(3)]
        void Remove(string strName);
    }

    //[ComImport]
    //[Guid("76A64169-CB41-11D1-8B02-00600806D9B6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ISWbemQualifier
    {
        [DispId(0)]
        string Name { get; }

        [DispId(1)]
        object Value { get; set; }

        [DispId(2)]
        bool IsLocal { get; }

        [DispId(3)]
        bool PropagatesToSubclass { get; }

        [DispId(4)]
        bool PropagatesToInstance { get; }
    }
}
