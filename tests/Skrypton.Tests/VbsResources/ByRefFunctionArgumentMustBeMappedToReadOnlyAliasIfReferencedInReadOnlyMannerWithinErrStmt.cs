
        public object F1(ref object a)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object a_zref = a;
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.CALLm1v0(this, a_zref ?? throw new InvalidOperationException("Reference not set:a_zref"), "Name"));
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
