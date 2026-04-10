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
            _outer.serv = VBScriptConstants.Nothing;
            if (_.IF(_.NOT(_.IS(_outer.serv, VBScriptConstants.Nothing))))
            {
                if (_.IF(_.EQ(_.CALLm1v1(this, _.NnO(_outer.serv, "serv"), "enabled", (Int16)7), true)))
                {
                    _.SETm1a1(this, _.NnO(_outer.serv, "serv"), "Enabled", (Int16)8, false);
                }
                else
                {
                    _.SETm1a1(this, _.NnO(_outer.serv, "serv"), "Enabled", (Int16)9, true);
                }
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
            serv = null;
        }
        internal object serv { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
    }
}
