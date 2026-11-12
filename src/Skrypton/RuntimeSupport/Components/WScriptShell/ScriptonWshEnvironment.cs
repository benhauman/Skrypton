using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// WshShell.Environment(type) - the variables of the type (issue #107): env(name) reads a variable ("" if there is none), env(name) = value
    /// sets it, Remove removes it (no error if there is none), For Each lists "NAME=value"
    /// </summary>
    [SourceClassName("IWshEnvironment")] // TypeName(CreateObject("WScript.Shell").Environment) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Item")]
#pragma warning disable CA1010 // Generic IEnumerable<T> - For Each over the "NAME=value" strings, as with the COM object
    internal sealed class ScriptonWshEnvironment : IEnumerable
#pragma warning restore CA1010
    {
        private readonly IHostEnvironmentVariablesHostService _variables;
        private readonly string _type;

        public ScriptonWshEnvironment(IHostEnvironmentVariablesHostService variables, string type)
        {
            _variables = variables ?? throw new ArgumentNullException(nameof(variables));
            _type = type ?? throw new ArgumentNullException(nameof(type));
        }

        [DispId(0)]
        [IsDefault]
        public string Item(string Name) => _variables.GetVariable(_type, RequireName(Name)) ?? "";

        [DispId(0)]
        [IsDefault]
        public void Item(string Name, string Value) => _variables.SetVariable(_type, RequireName(Name), Value ?? "");

        [DispId(1)]
        public int length => Count();

        [DispId(1610743809)]
        public int Count() => _variables.GetVariables(_type).Count;

        [DispId(1610743812)]
        public void Remove(string Name)
        {
            RequireName(Name);
            if (_variables.GetVariable(_type, Name) != null)
                _variables.SetVariable(_type, Name, null);
        }

        public IEnumerator GetEnumerator() => _variables.GetVariables(_type).Select(v => (object)(v.Key + "=" + v.Value)).ToList().GetEnumerator();

        [DispId(-4)] // DISPID_NEWENUM
        public IEnumerator _NewEnum() => GetEnumerator();

        private static string RequireName(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new InvalidProcedureCallOrArgumentException("'WshEnvironment' (a variable name is required)");
            return name;
        }
    }
}
