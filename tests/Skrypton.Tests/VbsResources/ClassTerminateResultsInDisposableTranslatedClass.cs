
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [SourceClassName(nameof(C1))]
    public sealed class C1 : IDisposable
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly EnvironmentReferences _env;
        private readonly GlobalReferences _outer;
        private bool _disposed;
        public C1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = outer ?? throw new ArgumentNullException(nameof(outer));
            _disposed = false;
        }
        ~C1()
        {
            try { Dispose(false); }
            catch(Exception e)
            {
                try { _.SETERROR(e); } catch { }
            }
        }
        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private void Dispose(bool disposing)
        {
            if (_disposed)
                return;
            if (disposing)
                Class_Terminate();
            _disposed = true;
        }
        public void Class_Terminate()
        {
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", "Gone!");
        }
    }
