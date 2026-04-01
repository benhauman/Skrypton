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
            _outer.objShell = _.OBJ(_.CREATEOBJECT("Shell.Application"));
            if (_.IF(_.NOTEQ(_.NullableSTR(_.TRIM(_.CALLm1v0(this, _env.TextBoxWebsite ?? throw new InvalidOperationException("Reference not set:TextBoxWebsite"), "Text"))), "")))
            {
                _.CALLm1v1(this, _outer.objShell ?? throw new InvalidOperationException("Reference not set:objShell"), "ShellExecute", _.CALLm1v0(this, _env.TextBoxWebsite ?? throw new InvalidOperationException("Reference not set:TextBoxWebsite"), "Text"));
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
            objShell = null;
        }
        internal object objShell { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object TextBoxWebsite { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
