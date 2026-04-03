using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [SourceClassName("WshShell")] // for TYPENAME(CreateObject("WScript.Shell"))
    //"WScript.Shell"; 'Windows Script Host Shell object'; WshShell COM object
    // CLSID: 72C24DD5-D70A-438B-8A42-98424B88AFB8, HKEY_CLASSES_ROOT\WScript.Shell, %SystemRoot%\System32\wshom.ocx -> Interop.IWshRuntimeLibrary.dll
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyWScriptShell : IReflectOnClrType, IWshShell3
    {
        private readonly IServiceProvider _hostServices;

        public MyWScriptShell(IServiceProvider hostServices) : base(typeof(MyWScriptShell))
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        [DispId(1000)]
        public int Run([In, MarshalAs(UnmanagedType.BStr)] string Command, [In, MarshalAs(UnmanagedType.Struct), DefaultParameterValue(1), Optional] object WindowStyle, [In, MarshalAs(UnmanagedType.Struct), Optional] object WaitOnReturn)
        {
            // WindowStyle:             0 = Hidden, 1 = Normal (DEFAULT), 3 = Maximized, etc.
            /*
               | Value | Meaning                 |
               | ----- | ----------------------- |
               | 0     | Hidden                  |
               | 1     | Normal (default)        |
               | 2     | Minimized               |
               | 3     | Maximized               |
               | 4     | Normal (no activate)    |
               | 7     | Minimized (no activate) |

             */
            // WaitOnReturn:    true = block until exit, false = Do not wait/non‑blocking (DEFAULT) // VARIANT_FALSE

            //System.Diagnostics.Process.Start("notepad.exe");
            var psi = new System.Diagnostics.ProcessStartInfo(Command);
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;   // default
            if (WindowStyle != null && WindowStyle is not System.Reflection.Missing)
            {
                long nWindowStyle = Convert.ToByte(WindowStyle);
                switch (nWindowStyle)
                {
                    case 0: psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden; break;
                    case 1: psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal; break;
                    case 2: psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized; break;
                    case 3: psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized; break;
                    case 4: /* Normal (no activate) */ break;
                    case 7: /* Minimized (no activate) */ break;
                }
            }

            bool bWaitOnReturn = false;
            if (WaitOnReturn != null && WaitOnReturn is not System.Reflection.Missing)
            {
                if (WaitOnReturn is bool boolWait)
                    bWaitOnReturn = boolWait;
                else
                    bWaitOnReturn = Convert.ToInt32(WaitOnReturn) != 0;
            }

            var svc = _hostServices.GetRequiredService<IHostProcessControlHostService>();
            svc.ProcessStart(Command, (byte)psi.WindowStyle, bWaitOnReturn);

            //throw new NotImplementedException($"[WshShell] Run(Command:{Command})");
            return 0; // S_OK;
        }

        public bool AppActivate([MarshalAs(UnmanagedType.Struct)] object App, [MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(null)] object Wait)
        {
            if (App is IReflect)
                return false;//MyOutlookApplicationClass
            int processId = Convert.ToInt32(App);
            var svc = _hostServices.GetRequiredService<IHostProcessControlHostService>();
            return svc.ProcessActivate(processId);// Process.GetProcessesByName("notepad"); SetForegroundWindow(processes[0].MainWindowHandle);
        }

        public object Exec(object strCommand)
        {
            // The Exec method returns a WshScriptExec object, which provides status and error information about a script run with Exec along with access to the StdIn, StdOut, and StdErr channels.The Exec method allows the execution of command line applications only. The Exec method cannot be used to run remote scripts. Do not confuse the Exec method with the Execute method (of the WshRemote object).
            return new MyWshScriptExec((string)strCommand);
        }
    }

    [SourceClassName("WshScriptExec")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyWshScriptExec : IReflectOnClrType
    {
        public string CommandText { get; }

        public MyWshScriptExec(string commandText)
        {
            CommandText = commandText;
        }
    }

    internal static class IHostProcessControlHostServiceExtensions
    {
        public static T GetRequiredService<T>(this IServiceProvider serviceProvider) where T : class
        {
            if (serviceProvider == null) throw new ArgumentNullException(nameof(serviceProvider));
            return ((T)serviceProvider.GetService(typeof(T))) ?? throw new InvalidOperationException($"Service '{typeof(T).FullName}' not registered.");
        }
    }

    public interface IHostProcessControlHostService
    {
        void ProcessStart(string command, byte windowMode, bool waitOnReturn);
        bool ProcessActivate(int processId);
        void ProcessesCollect(Func<int, string, bool> collector);
    }

    internal interface IWshShell
    {
        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1000)]
        int Run([In][MarshalAs(UnmanagedType.BStr)] string Command,
            [Optional][In][MarshalAs(UnmanagedType.Struct), DefaultParameterValue(1), DefaultValue(2)] object WindowStyle,
            [Optional][In][MarshalAs(UnmanagedType.Struct), DefaultParameterValue(false)] object WaitOnReturn);

        [DispId(1000)]
        bool AppActivate(
            [MarshalAs(UnmanagedType.Struct)] object App,
            [Optional][MarshalAs(UnmanagedType.Struct), DefaultParameterValue(null)] object Wait
        );
    }

    internal interface IWshShell2 : IWshShell
    {
    }

    internal interface IWshShell3 : IWshShell2
    { }
}
