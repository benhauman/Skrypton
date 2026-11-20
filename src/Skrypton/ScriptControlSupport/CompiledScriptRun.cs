using System;
using Skrypton.RuntimeSupport;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// One run of a CompiledScript (see issue #187): its outer scope has run and its procedures can be called by name. It must be disposed when
    /// the request is done (it owns the compat layer of the run).
    /// </summary>
    internal sealed class CompiledScriptRun : IDisposable
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _compatLayer;
        private readonly GlobalReferencesBase _globalReferences;

        public CompiledScriptRun(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, GlobalReferencesBase globalReferences)
        {
            _compatLayer = compatLayer;
            _globalReferences = globalReferences;
        }

        /// <summary>
        /// Call a Sub or Function of the script by name through the VBScript binder, as the script itself calls it: the arguments are passed
        /// as VBScript passes them (a ByRef parameter that the procedure changes updates its element of the arguments array)
        /// </summary>
        public object? Call(string procedureName, object?[] arguments)
        {
            IBuildCallArgumentProviders argumentProviders = _compatLayer.ARGS;
            for (int i = 0; i < arguments.Length; i++)
            {
                int index = i;
                argumentProviders = argumentProviders.Ref(arguments[i]!, value => arguments[index] = value);
            }
            return _compatLayer.CALL(_globalReferences, _globalReferences, new[] { procedureName }, argumentProviders.GetArgs());
        }

        public void Dispose() => _compatLayer.Dispose();
    }
}
