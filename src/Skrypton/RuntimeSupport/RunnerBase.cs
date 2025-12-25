using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace Skrypton.RuntimeSupport
{
    public abstract class RunnerBase
    {
        public RunnerBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            if (compatLayer == null) throw new ArgumentNullException(nameof(compatLayer));
        }

        public abstract EnvironmentReferencesBase CreateEnvironmentReferencesInstance();
    }

    public abstract class RunnerBaseT<TEnvironmentReferences, TGlobalReferencesBase> : RunnerBase
        where TEnvironmentReferences : EnvironmentReferencesBase, new()
        where TGlobalReferencesBase : GlobalReferencesBase
    {
        public RunnerBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
        }
        public override EnvironmentReferencesBase CreateEnvironmentReferencesInstance()
        {
            return new TEnvironmentReferences();
        }
    }

    public abstract class EnvironmentReferencesBase
    {
        private readonly Dictionary<string, object> _externalReferences = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        protected EnvironmentReferencesBase()
        {

        }

        public void InitializeExternalReference(string referenceName, object reference)
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            _externalReferences[referenceName] = reference ?? throw new ArgumentNullException(nameof(reference)); // Use DBValue.Null for nulls.
        }

        protected object GetExternalReferenceAsObject([CallerMemberName] string referenceName = "")
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            if (_externalReferences.TryGetValue(referenceName, out object reference))
                return reference;
            return null;//?!?
        }
        protected void RestoreExternalReferenceAsObject(object newInstance, [CallerMemberName] string referenceName = "")
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            var current = GetExternalReferenceAsObject(referenceName);
            if (current != newInstance)
                throw new InvalidOperationException("not same");
        }
    }
    public abstract class GlobalReferencesBase
    {
        protected GlobalReferencesBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferencesBase env)
        {

        }
    }
    public abstract class GlobalReferencesBaseT<TEnvironmentReferences> : GlobalReferencesBase where TEnvironmentReferences : EnvironmentReferencesBase
    {
        protected GlobalReferencesBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences env)
            : base(compatLayer, env)
        {

        }
    }
}
