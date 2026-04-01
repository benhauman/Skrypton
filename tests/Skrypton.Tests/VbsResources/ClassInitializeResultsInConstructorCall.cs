
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
            try { Class_Initialize(); }
            catch(Exception e)
            {
                _.SETERROR(e);
            }
        }
        public void Class_Initialize()
        {
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Here!");
        }
    }
