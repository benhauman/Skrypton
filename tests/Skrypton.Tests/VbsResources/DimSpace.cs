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
            _outer.Space = "+";
            _outer.URLEncode = _.CONCAT(_.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"), "x", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"));
            _outer.URLEncode = _.CONCAT("y", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"));
            _outer.URLEncode = _.CONCAT(_.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"), "z");
            _outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F1", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE")));
            _outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F2", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE")));
            _outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F3", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE")));
            _outer.URLEncode = _.VAL(_.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F4", _.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE")));
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
            i = null;
            CharCode = null;
            rewritten_Char = null;
            Space = null;
            URLEncode = null;
        }
        internal object i { get; set; }
        internal object CharCode { get; set; }
        internal object rewritten_Char { get; set; }
        internal object Space { get; set; }
        internal object URLEncode { get; set; }
        public object F1()
        {
            object F1_retVal = null;
            object Space = null;
            return F1_retVal;
        }
        public object F2(ref object Space)
        {
            return _.VAL(_.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"));
        }
        public object F3(ref object Space)
        {
            return _.VAL(_.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"));
        }
        public object F4(object Space)
        {
            return _.VAL(_.CALLm1v0(this, _ ?? throw new InvalidOperationException("Reference not set:_"), "SPACE"));
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
    }
}
