using System;
using System.Collections.Generic;
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
    internal sealed class MyWScriptShell : IDispatchBase.IReflectOnClrType, IWshShell3
    {
        private readonly IServiceProvider _hostServices;

        public MyWScriptShell(IServiceProvider hostServices) : base(typeof(MyWScriptShell))
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        [DispId(1000)]
        public int Run([In, MarshalAs(UnmanagedType.BStr)] string Command, [In, MarshalAs(UnmanagedType.Struct), Optional] object WindowStyle, [In, MarshalAs(UnmanagedType.Struct), Optional] object WaitOnReturn)
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
    }

    internal interface IWshShell
    {
        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1000)]
        int Run([In][MarshalAs(UnmanagedType.BStr)] string Command, [Optional][In][MarshalAs(UnmanagedType.Struct)] object WindowStyle, [Optional][In][MarshalAs(UnmanagedType.Struct)] object WaitOnReturn);

    }

    internal interface IWshShell2 : IWshShell
    {
    }

    internal interface IWshShell3 : IWshShell2
    { }
}
