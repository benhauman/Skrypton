using System;
using System.Collections;
using Skrypton.RuntimeSupport;
namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }
        protected override GlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) => new GlobalReferences(compatLayer, env);
        protected override void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences globalReferences)
        {
            var _env = env ?? throw new ArgumentNullException(nameof(env));
            var _outer = globalReferences ?? throw new ArgumentNullException(nameof(globalReferences));

            if (_.IF(_.EQ(_.NUM(_outer.priority), (Int16)1)))
            {
                _.CALLm1v5(this, _env.hlObj, "SetValue", "CaseGeneral.Priority", (Int16)101, (Int16)102, (Int16)103, "PriorityNormal");
            }
            else if (_.IF(_.EQ(_.NUM(_outer.priority), (Int16)2)))
            {
                _.CALLm1v5(this, _env.hlObj, "SetValue", "CaseGeneral.Priority", (Int16)201, (Int16)202, (Int16)203, "PriorityMedium");
            }
            else
            {
                _.CALLm1v5(this, _env.hlObj, "SetValue", "CaseGeneral.Priority", (Int16)901, (Int16)902, (Int16)903, "");
            }
        }
    }
    public sealed class GlobalReferences : GlobalReferencesBaseT<EnvironmentReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly GlobalReferences _outer;
        private readonly EnvironmentReferences _env;
        public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) : base(compatLayer, env)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = this;
            priority = null;
        }
        internal object priority { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlObj { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
