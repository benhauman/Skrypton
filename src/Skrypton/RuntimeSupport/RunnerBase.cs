using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace Skrypton.RuntimeSupport
{
    public abstract class RunnerBase
    {
        internal IProvideVBScriptCompatFunctionalityToIndividualRequests CompatLayer { get; } // rename it to '_'
        protected RunnerBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            CompatLayer = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }

        public abstract EnvironmentReferencesBase CreateEnvironmentReferencesInstance();

        public static RunnerBase CreateRunnerInstanceForType(Type runnerType, IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            return (RunnerBase)Activator.CreateInstance(runnerType, [compatLayer]);
        }

        public abstract GlobalReferencesBase Run(EnvironmentReferencesBase environmentReferences);
    }

    public abstract class RunnerBaseT<TEnvironmentReferences, TGlobalReferences> : RunnerBase
        where TEnvironmentReferences : EnvironmentReferencesBase, new()
        where TGlobalReferences : GlobalReferencesBase
    {
        protected RunnerBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
        }
        public override EnvironmentReferencesBase CreateEnvironmentReferencesInstance()
        {
            return new TEnvironmentReferences();
        }

        protected abstract TGlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences environmentReferences);

        public override GlobalReferencesBase Run(EnvironmentReferencesBase environmentReferences)
        {
            TGlobalReferences globalReferences = CreateGlobalReferences(CompatLayer, (TEnvironmentReferences)environmentReferences);
            Go(CompatLayer, (TEnvironmentReferences)environmentReferences, globalReferences);
            return globalReferences;
        }

        protected abstract void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, TEnvironmentReferences environmentReferences, TGlobalReferences globalReferences);
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
            object current = GetExternalReferenceAsObject(referenceName);
            if (current != null && newInstance != null && current != newInstance)
            {
                throw new InvalidOperationException("not same");
            }
            else
            {
                if (newInstance == null)
                    _externalReferences.Remove(referenceName);
                else
                    _externalReferences[referenceName] = newInstance;
            }
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
