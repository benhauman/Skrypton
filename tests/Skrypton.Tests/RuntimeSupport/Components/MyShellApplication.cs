using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /*
     * // Define the COM-visible interface
        [Guid("EAA4976A-45C3-4BC5-BC0B-E474F4C3C83F")]
       [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
       public interface IMyShell
       {
           [DispId(0)]   // Default member
            string DefaultMember { get; }

           [DispId(1)]
           void Open(string path);

           [DispId(2)]
           string GetVersion();
       }


     */

    [SourceClassName("Shell")] // for TYPENAME(CreateObject("Shell.Application"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[DefaultMember("Item")]
    internal sealed class MyShellApplication
    {
/*
- Type library:
   - Shell.Application is defined in the Shell type library (Shell32.dll).
   - You can inspect it with tools like OLE/COM Object Viewer (oleview.exe) or tlbimp.exe (Type Library Importer).
   - ProgID: "Shell.Application" → CLSID {13709620-C279-11CE-A49E-444553540000}.
   - IDL / Header files:
   - The automation interfaces (IShellDispatch, IShellDispatch2, … up to IShellDispatch6) are declared in the Windows SDK headers:
   - <shldisp.h> → contains the COM interface definitions for Shell Dispatch objects.
   - <exdisp.h> → contains Explorer-related COM interfaces.
   - These headers are included with the Windows SDK (e.g., in C:\Program Files (x86)\Windows Kits\10\Include\<version>\um\).
 */

        // - It’s a fire‑and‑forget call — it launches the program or document but doesn’t give you a process handle or exit code.
        public void ShellExecute(
            object file        // Required: the file or program to run
        //Arguments,   ' Optional: command-line arguments
        //Directory,   ' Optional: working directory
        //Operation,   ' Optional: verb (e.g. "open", "print")
        //Show         ' Optional: window style (SW_SHOW, SW_HIDE, etc.)

            )
        {
            /*
            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
               static extern IntPtr ShellExecute(
                   IntPtr hwnd,
                   string lpOperation,
                   string lpFile,
                   string lpParameters,
                   string lpDirectory,
                   int nShowCmd);

            or
            var psi = new ProcessStartInfo
               {
                   FileName = "example.txt",
                   Verb = "print",   // other verbs: "open", "edit"
                UseShellExecute = true
               };
               Process.Start(psi);

             */
        }
    }

}