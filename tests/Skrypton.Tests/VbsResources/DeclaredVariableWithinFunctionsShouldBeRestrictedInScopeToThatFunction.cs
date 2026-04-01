
            _.CALLm1v0(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "Test1");
        public object Test1()
        {
            object Test1_retVal = null;
            object i = null;
            _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.ARGS.Ref(i, v => { i = v; }));
            return Test1_retVal;
        }
