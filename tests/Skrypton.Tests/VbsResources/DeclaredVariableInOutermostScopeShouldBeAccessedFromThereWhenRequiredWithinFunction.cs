
            _.CALLm1v0(this, _outer, "Test1");
        public object Test1()
        {
            object Test1_retVal = null;
            _.CALLm1argp(this, _.NnO(_env.WScript, "WScript"), "Echo", _.ARGS.Ref(_outer.i, v => { _outer.i = v; }));
            return Test1_retVal;
        }
