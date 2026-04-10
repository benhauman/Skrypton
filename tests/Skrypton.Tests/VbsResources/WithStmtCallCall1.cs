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
            var with = _.OBJ(_outer.adoSQLCmdParam);
            _.SETm1a0(this, _.NnO(with, "with"), "ActiveConnection", VBScriptConstants.Nothing);
            _.CALLm2v1(this, _.NnO(with, "with"), "Pr", "Ap", _.CALLm1v3(this, _.NnO(with, "with"), "CreateParameterX", "RETURN_VALUEx", (Int16)3, (Int16)4));
            _.CALLm2v1(this, _.NnO(with, "with"), "Parameters", "Append", _.CALLm1v5(this, _.NnO(with, "with"), "CreateParameterY", "@FirstCharName", (Int16)202, (Int16)1, (Int16)1, "FirstCharName"));
            _.CALLm1v0(this, _.NnO(with, "with"), "Execute");
            _outer.parmval = _.VAL(_.CALLm1v0(this, _.NnO(_.CALLm1v1(this, _.NnO(with, "with"), "Parameters", (Int16)2), "(_.call result)"), "Value"));
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
            parmval = null;
            adoSQLCmdParam = null;
        }
        internal object parmval { get; set; }
        internal object adoSQLCmdParam { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
