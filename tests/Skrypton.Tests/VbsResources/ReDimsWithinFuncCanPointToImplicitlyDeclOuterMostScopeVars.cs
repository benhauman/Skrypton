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

            _env.a = (Int16)1;
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

        public object F1()
        {
            object F1_retVal = null;
            _env.a = _.NEWARRAY(new object[] { (Int16)2 }); // This refers to the implicitly-declared variable "a" in the outermost scope
            return F1_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object a { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [SourceClassName(nameof(C1))]
    public sealed class C1
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly EnvironmentReferences _env;
        private readonly GlobalReferences _outer;
        public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = outer ?? throw new ArgumentNullException(nameof(outer));
            c = null;
        }

        private object c { get; set; }

        public object CF1()
        {
            object CF1_retVal = null;
            object b = null;
            _env.a = _.NEWARRAY(new object[] { (Int16)3 }); // This refers to the implicitly-declared variable "a" in the outermost scope
            b = _.NEWARRAY(new object[] { (Int16)3 }); // There is no reference for this to relate to, so it acts as new explicit variable declaration
            c = _.NEWARRAY(new object[] { (Int16)3 }); // This refers to the explicitly-declared variable "c" in the containing class
            return CF1_retVal;
        }
    }
}