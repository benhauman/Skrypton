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
            _outer.status = "z";
            if (_.IF(_.EQ(_outer.status, "SRMInternalStateOpen")))
            {
                _outer.internalstatus = "OPEN";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateToBeChecked")))
            {
                _outer.internalstatus = "TOBECHECKED";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateWaitingForAnalyse")))
            {
                _outer.internalstatus = "WAITING_FOR";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateWaitingForOffer")))
            {
                _outer.internalstatus = "WAITING_FOR";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateOrderReceived")))
            {
                _outer.internalstatus = "TOBECHECKED";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateSolved")))
            {
                _outer.internalstatus = "SOLVED";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateClosed")))
            {
                _outer.internalstatus = "CLOSED";
            }
            else if (_.IF(_.EQ(_outer.status, "SRMInternalStateWaitingForInternal")))
            {
                _outer.internalstatus = "WAITING_FOR";
            }
            else
            {
                _outer.internalstatus = "OPEN";
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
            status = null;
            internalstatus = null;
        }
        internal object status { get; set; }
        internal object internalstatus { get; set; }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
