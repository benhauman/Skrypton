            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v0(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "Func1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
        public object Func1()
        {
            object Func1_retVal = null;
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", "Test1");
            return Func1_retVal;
        }
