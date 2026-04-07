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
            _outer.searchResult = _.VAL(_.CALLm1argp(this, _, "ARRAY", _.ARGS.ForceBrackets())); // Initialize as empty array

            if (_.IF(_.AND(_.GTE(_.NullableNUM(_.UBOUND(_outer.searchResult)), (Int16)0), _.NOTEQ(_.NullableSTR(_.CALLm0argp(this, _.CALLm0argp(this, _outer.searchResult ?? throw new InvalidOperationException("Reference not set:searchResult"), _.ARGS.Val((Int16)0)) ?? throw new InvalidOperationException("Reference not set:(_.call result)"), _.ARGS.Val((Int16)2))), ""))))
            {
                _.MSGBOX("T");
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
            searchResult = null;
        }
        internal object searchResult { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
    }
}
