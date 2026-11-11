// Moved from the test project as-is; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts (so they must stay instance members) and the names mirror the COM type library
#pragma warning disable CA1822, CA1008, CA1711, CA1716, CA1309, CA1859
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.RuntimeSupport.Components
{
    [SourceClassName("IGenScriptletTLib")] // TypeName(CreateObject("Scriptlet.TypeLib")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]

    internal sealed class ScriptonScriptletTypeLib : ScriptonIReflectOnClrType
    {
        public ScriptonScriptletTypeLib()
        {
        }

        /// <summary>
        /// A new GUID in registry format, eg. "{3F2504E0-4F89-11D3-9A0C-0305E82C3301}" - as with Windows, the value is 40 characters long: the
        /// 38 characters of the GUID followed by two NUL characters (scripts use Left(.., 38) or Mid(.., 2, 36))
        /// </summary>
        [DispId(1)] // not the default member: VarType(CreateObject("Scriptlet.TypeLib")) is 9 (verified with cscript)
        public string GUID => "{" + System.Guid.NewGuid().ToString().ToUpperInvariant() + "}" + "\0\0";
    }

    /*
[
        uuid(06290BD3-48AA-11D2-8432-006008C3FBFC),
        dual,
        oleautomation
    ]
    interface IScriptletTypeLib : IDispatch
    {
        [propget, id(0)]
        HRESULT GUID([out, retval] BSTR* pVal);
    };
     */
}