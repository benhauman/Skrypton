using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [SourceClassName("TypeLib")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember(nameof(MyScriptletTypeLib.GUID))]  // +[DispId(0)] +[IsDefault] // "Scriptlet.TypeLib"

    internal sealed class MyScriptletTypeLib : IReflectOnClrType
    {
        public MyScriptletTypeLib()
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