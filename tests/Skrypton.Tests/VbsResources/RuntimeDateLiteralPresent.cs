using System;
using System.Collections;
using System.Collections.ObjectModel;
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
            RuntimeDateLiteralValidator.ValidateAgainstCurrentCulture(_);

            if (_.IF(_.EQ(_.NullableDATE(_env.a), _.DateLiteralParser.Parse("29 May 2015"))))
            {
            }
        }
        private static class RuntimeDateLiteralValidator
        {
            private static readonly ReadOnlyCollection<Tuple<string, int[]>> _literalsToValidate =
            new ReadOnlyCollection<Tuple<string, int[]>>(new[] {
                Tuple.Create("29 May 2015", new[] { 2 })
            });

            public static void ValidateAgainstCurrentCulture(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
            {
                if (compatLayer == null)
                    throw new ArgumentNullException(nameof(compatLayer));
                foreach (var dateLiteralValueAndLineNumbers in _literalsToValidate)
                {
                    try { compatLayer.DateLiteralParser.Parse(dateLiteralValueAndLineNumbers.Item1); }
                    catch
                    {
                        throw new SyntaxError(string.Format(
                            "Invalid date literal #{0}# on line{1} {2}",
                            dateLiteralValueAndLineNumbers.Item1,
                            (dateLiteralValueAndLineNumbers.Item2.Length == 1) ? "" : "s",
                            string.Join<int>(", ", dateLiteralValueAndLineNumbers.Item2)
                        ));
                    }
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
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object a { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
