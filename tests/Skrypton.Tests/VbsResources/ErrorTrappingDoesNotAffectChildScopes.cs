            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v0(this, _outer, "Func1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
        public object Func1()
        {
            object Func1_retVal = null;
            _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Test1");
            return Func1_retVal;
        }
