
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
        }
        public object Go()
        {
            object Go_retVal = null;
            object a = null; /* Undeclared in source */
            a = _.OBJ(_.CALLm2v0(this, _.NnO(this, "this"), "GetSomething", "Name"));
            return Go_retVal;
        }
        public object GetSomething()
        {
            return null;
        }
    }
