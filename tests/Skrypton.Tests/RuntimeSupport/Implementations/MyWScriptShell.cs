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
    internal sealed class MyWScriptShell : IWshShell3
    {
        public MyWScriptShell()
        {

        }

        [DispId(1000)]
        public int Run([In, MarshalAs(UnmanagedType.BStr)] string Command, [In, MarshalAs(UnmanagedType.Struct), Optional] ref object WindowStyle, [In, MarshalAs(UnmanagedType.Struct), Optional] ref object WaitOnReturn)
        {
            throw new NotImplementedException($"[WshShell] Run(Command:{Command})");
        }
    }

    internal interface IWshShell
    {
        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1000)]
        int Run([In][MarshalAs(UnmanagedType.BStr)] string Command, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WindowStyle, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object WaitOnReturn);

    }

    internal interface IWshShell2 : IWshShell
    {
    }

    internal interface IWshShell3 : IWshShell2
    { }
}
