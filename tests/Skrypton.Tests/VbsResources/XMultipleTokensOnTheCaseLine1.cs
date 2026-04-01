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
        }
        public void PriorityMatrix()
        {
            object impact = null;
            object urgency = null;
            object impactText = null;
            object urgencyText = null;
            object priority = null;
            object priorityText = null;
            object hlObj = null; /* Undeclared in source */
            object ComboBoxImpact = null; /* Undeclared in source */
            impactText = _.VAL(_.CALLm1v5(this, hlObj ?? throw new InvalidOperationException("Reference not set:hlObj"), "GetValue", "IncidentAttribute.Impact", (Int16)0, (Int16)0, (Int16)0, (Int16)0));

            if (_.IF(_.EQ(impactText, "ImpactSinglePerson")))
            {
                impact = (Int16)1;
            }
            else if (_.IF(_.EQ(impactText, "ImpactMultipleGroups")))
            {
                impact = (Int16)2;
            }
            else if (_.IF(_.EQ(impactText, "ImpactEntireOrganization")))
            {
                impact = (Int16)3;
            }
            else if (_.IF(_.EQ(impactText, "")))
            {
                impact = (Int16)0;
            }
            else
            {
                impact = _.VAL(_.CALLm1argp(this, ComboBoxImpact ?? throw new InvalidOperationException("Reference not set:ComboBoxImpact"), "GetCurSel", _.ARGS.ForceBrackets()));
            }

        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
    }
}
