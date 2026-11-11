// Moved from the test project as-is; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts (so they must stay instance members) and the names mirror the COM type library
#pragma warning disable CA1822, CA1008, CA1711, CA1716, CA1309, CA1859
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.RuntimeSupport.Components
{
    [SourceClassName("TypeLib")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember(nameof(ScriptonScriptletTypeLib.GUID))]  // +[DispId(0)] +[IsDefault] // "Scriptlet.TypeLib"

    internal sealed class ScriptonScriptletTypeLib : ScriptonIReflectOnClrType
    {
        public ScriptonScriptletTypeLib()
        {
        }

        [DispId(0)]
        public string GUID()
        {
            return System.Guid.NewGuid().ToString();
        }
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