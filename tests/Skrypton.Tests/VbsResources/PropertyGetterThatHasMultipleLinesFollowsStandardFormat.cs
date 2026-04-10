
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
        [TranslatedProperty("Name")]
        public object Name()
        {
            object Name_retVal = null;
            _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "get_Name");
            Name_retVal = "C1";
            return Name_retVal;
        }
    }
