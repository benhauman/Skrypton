using System;
using System.Collections;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Compat;

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
            _.SET(VBScriptConstants.Nothing, this, with, "ActiveConnection");
            _.CALLm2argp(this, with, "Pr", "Ap", _.ARGS.Val(_.CALLm1argp(this, with, "CreateParameterX", _.ARGS.Val("RETURN_VALUEx").Val((Int16)3).Val((Int16)4))));
            _.CALLm2argp(this, with, "Parameters", "Append", _.ARGS.Val(_.CALLm1argp(this, with, "CreateParameterY", _.ARGS.Val("@FirstCharName").Val((Int16)202).Val((Int16)1).Val((Int16)1).Val("FirstCharName"))));
            _.CALLm1v0(this, with, "Execute");
            _outer.parmval = _.VAL(_.CALLm1v0(this, _.CALLm1argp(this, with, "Parameters", _.ARGS.Val((Int16)2)), "Value"));
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
