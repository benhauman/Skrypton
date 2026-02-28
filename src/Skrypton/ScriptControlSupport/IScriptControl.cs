using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Skrypton.ScriptControlSupport
{
/*
   | Property                                              | Type                  | Default Value     | Setter Effect / Notes                                                                                           | Typical Use Cases                                             |
   | ----------------------------------------------------- | --------------------- | ----------------- | --------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------- |
   | `Language`                                            | string                | `""` (empty)      | Must set before running scripts; `"VBScript"` or `"JScript"`                                                    | Select scripting language for code execution                  |
   | `AllowUI`                                             | bool                  | `false`           | If `true`, scripts can show MsgBox/InputBox/etc.                                                                | Enable UI interaction from scripts                            |
   | `Timeout`                                             | int                   | `-1`              | Milliseconds; `-1` = infinite                                                                                   | Limit script execution time to prevent hangs                  |
   | `State`                                               | `ScriptControlStates` | `Initialized (0)` | 0 → reset engine (clears code, variables, modules)<br>1 → engine connected / ready to run code & receive events | Control engine lifecycle; reset or pre-start                  |
   | `Modules`                                             | `ScriptModules`       | empty collection  | Read-only; adding code implicitly adds to default module                                                        | Inspect all loaded modules/functions; usually 1 module exists |
   | `UseSafeSubset`                                       | bool                  | `true`            | `false` = allows full scripting access, potentially unsafe                                                      | Security measure to restrict dangerous operations             |
   | `SitehWnd`                                            | int                   | `0`               | HWND of host window; required if AllowUI = true                                                                 | Needed for script dialogs/MsgBox                              |
   | `Error`                                               | `ScriptError`         | `null/Nothing`    | Read-only; reflects **last runtime error**                                                                      | Check for runtime errors after `Run` or `ExecuteStatement`    |
   | `CodeObject`                                          | object                | `null/Nothing`    | Read-only; dynamic interface to **all global functions/subs/variables**                                         | Call functions or access global variables programmatically    |
   | `AddCode(string code)`                                | method                | n/a               | Adds VBScript/JScript code to default module                                                                    | Define functions, subs, variables                             |
   | `Run(string functionName, params object[] args)`      | method                | n/a               | Calls a function in the default module                                                                          | Execute script functions and get return value                 |
   | `ExecuteStatement(string statement)`                  | method                | n/a               | Executes a statement or subroutine                                                                              | Run code without return value                                 |
   | `AddObject(string name, object obj, bool addMembers)` | method                | n/a               | Expose host object to the script                                                                                | Allow scripts to call back into host app                      |
 */
/*
   | Action                           | State Before    | State After     | Effect                                             |
   | -------------------------------- | --------------- | --------------- | -------------------------------------------------- |
   | Create ScriptControl             | N/A             | Initialized (0) | Engine created, empty                              |
   | AddCode / Run / ExecuteStatement | Initialized (0) | Connected (1)   | Engine starts executing code                       |
   | Set `State = Initialized`        | Connected (1)   | Initialized (0) | Resets engine; clears all code, variables, modules |
   | Set `State = Connected`          | Initialized (0) | Connected (1)   | Pre-starts engine; makes it ready to run code      |
 */
    //[ComImport]
    //[Guid("0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC")]
    //[TypeLibType(4304)] // 4304 = TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FNonExtensible | TypeLibTypeFlags.FDual | TypeLibTypeFlags.FHidden
    public interface IScriptControl // .\ThirdParty\interop\Interop.MSScriptControl.dll
    {
        [DispId(1500)]
        string Language
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1500)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1500)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [DispId(1501)]
        ScriptControlStates State
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [TypeLibFunc(1024)]
            [DispId(1501)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1501)]
            [TypeLibFunc(1024)]
            [param: In]
            set;
        }

        [DispId(1502)]
        int SitehWnd
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [TypeLibFunc(1024)]
            [DispId(1502)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1502)]
            [TypeLibFunc(1024)]
            [param: In]
            set;
        }

        [DispId(1503)]
        int Timeout
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1503)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1503)]
            [param: In]
            set;
        }

        [DispId(1504)]
        bool AllowUI
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1504)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1504)]
            [param: In]
            set;
        }

        [DispId(1505)]
        bool UseSafeSubset
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1505)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1505)]
            [param: In]
            set;
        }

        [DispId(1506)]
        Modules Modules
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1506)]
            [TypeLibFunc(1024)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1507)]
        Error Error
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1507)]
            [TypeLibFunc(1024)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1000)]
        object CodeObject
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1000)]
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
        }

        [DispId(1001)]
        Procedures Procedures
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1001)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(-552)]
        [TypeLibFunc(64)]
        void _AboutBox();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2500)]
        void AddObject([In][MarshalAs(UnmanagedType.BStr)] string Name, [In][MarshalAs(UnmanagedType.IDispatch)] object Object, [In] bool AddMembers = false);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2501)]
        void Reset();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2000)]
        void AddCode([In][MarshalAs(UnmanagedType.BStr)] string Code);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2001)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object Eval([In][MarshalAs(UnmanagedType.BStr)] string Expression);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2002)]
        void ExecuteStatement([In][MarshalAs(UnmanagedType.BStr)] string Statement);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2003)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object Run([In][MarshalAs(UnmanagedType.BStr)] string ProcedureName, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] Parameters);
    }
    public enum ScriptControlStates
    {
        Initialized,
        Connected
    }
    //[ComImport]
    //[TypeLibType(4304)]
    //[Guid("70841C78-067D-11D0-95D8-00A02463AB28")]
    public interface IScriptError
    {
        [DispId(201)]
        int Number
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(201)]
            get;
        }

        [DispId(202)]
        string Source
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(202)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(203)]
        string Description
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(203)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(204)]
        string HelpFile
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(204)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(205)]
        int HelpContext
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(205)]
            get;
        }

        [DispId(-517)]
        string Text
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(-517)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(206)]
        int Line
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(206)]
            get;
        }

        [DispId(-529)]
        int Column
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(-529)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(208)]
        void Clear();
    }
    //[ComImport]
    //[Guid("70841C78-067D-11D0-95D8-00A02463AB28")]
    //[CoClass(typeof(ErrorClass))]
    public interface Error : IScriptError
    {
    }
    //[ComImport]
    //[Guid("70841C73-067D-11D0-95D8-00A02463AB28")]
    [DefaultMember("Name")]
    //[TypeLibType(4304)]
    public interface IScriptProcedure
    {
        [DispId(0)]
        string Name
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(100)]
        int NumArgs
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(100)]
            get;
        }

        [DispId(101)]
        bool HasReturnValue
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(101)]
            get;
        }
    }
    //[ComImport]
    //[CoClass(typeof(ProcedureClass))]
    //[Guid("70841C73-067D-11D0-95D8-00A02463AB28")]
    public interface Procedure : IScriptProcedure
    {
    }
    //[ComImport]
    //[TypeLibType(4304)]
    //[Guid("70841C71-067D-11D0-95D8-00A02463AB28")]
    public interface IScriptProcedureCollection : System.Collections.IEnumerable, IEnumerable<Procedure>
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [TypeLibFunc(64)]
        [DispId(-4)]
        [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        new System.Collections.IEnumerator GetEnumerator();

        [DispId(0)]
        Procedure this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1)]
        int Count
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            get;
        }
    }
    //[ComImport]
    //[Guid("70841C71-067D-11D0-95D8-00A02463AB28")]
    //[CoClass(typeof(ProceduresClass))]
    public interface Procedures : IScriptProcedureCollection
    {
    }

    //[ComImport]
    [DefaultMember("Name")]
    //[TypeLibType(4304)]
    //[Guid("70841C70-067D-11D0-95D8-00A02463AB28")]
    public interface IScriptModule
    {
        [DispId(0)]
        string Name
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
        }

        [DispId(1000)]
        object CodeObject
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1000)]
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
        }

        [DispId(1001)]
        Procedures Procedures
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1001)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2000)]
        void AddCode([In][MarshalAs(UnmanagedType.BStr)] string Code);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2001)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object Eval([In][MarshalAs(UnmanagedType.BStr)] string Expression);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2002)]
        void ExecuteStatement([In][MarshalAs(UnmanagedType.BStr)] string Statement);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2003)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object Run([In][MarshalAs(UnmanagedType.BStr)] string ProcedureName, [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] Parameters);
    }
    //[ComImport]
    //[Guid("70841C70-067D-11D0-95D8-00A02463AB28")]
    //[CoClass(typeof(ModuleClass))]
    public interface Module : IScriptModule
    {
    }
    //[ComImport]
    //[TypeLibType(4304)]
    //[Guid("70841C6F-067D-11D0-95D8-00A02463AB28")]
    public interface IScriptModuleCollection : System.Collections.IEnumerable, IEnumerable<Module>
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(-4)]
        [TypeLibFunc(64)]
        [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        new System.Collections.IEnumerator GetEnumerator();

        [DispId(0)]
        Module this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1)]
        int Count
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(2)]
        [return: MarshalAs(UnmanagedType.Interface)]
        Module Add([In][MarshalAs(UnmanagedType.BStr)] string Name, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Object);
    }
    //[ComImport(t)]
    //[Guid("70841C6F-067D-11D0-95D8-00A02463AB28")]
    //[CoClass(typeof(ModulesClass))]
    public interface Modules : IScriptModuleCollection
    {
    }
    internal sealed class TypeLibFuncAttribute : System.Attribute
    {
        public TypeLibFuncAttribute(int flags) { }
    }
    //internal sealed class TypeLibTypeAttribute : System.Attribute
    //{
    //    public TypeLibTypeAttribute(int flags) { }
    //}

    //[Guid("AE311340-082C-11D0-95DE-00A02463AB28")]
    public abstract class ScriptControlConstants
    {
        [MarshalAs(UnmanagedType.LPStr)]
        public const string GlobalModule = "Global";

        public const int NoTimeout = -1;
    }
}