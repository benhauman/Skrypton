
            _.CALLm1v0(this, _outer, "Func1");
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", "Test2");
        public object Func1()
        {
            object Func1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", "Test1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return Func1_retVal;
        }
