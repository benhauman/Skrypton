using System;
using System.Collections.Generic;
using System.Text;

namespace Skrypton.RuntimeSupport
{
    public abstract class RunnerBase
    {
        public RunnerBase(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            if (compatLayer == null) throw new ArgumentNullException(nameof(compatLayer));
        }
    }

    public abstract class RunnerBaseT<TEnvironmentReferences, TGlobalReferencesBase> : RunnerBase
        where TEnvironmentReferences : EnvironmentReferencesBase, new()
        where TGlobalReferencesBase : GlobalReferencesBase
    {
        public RunnerBaseT(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
        }
    }

    public abstract class EnvironmentReferencesBase
    {
        protected EnvironmentReferencesBase()
        {

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
