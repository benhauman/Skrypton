
        public object F1(ref object a)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object a_zref = a;
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", _.CALLm1v0(this, _.NnO(a_zref, "a_zref"), "Name"));
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
