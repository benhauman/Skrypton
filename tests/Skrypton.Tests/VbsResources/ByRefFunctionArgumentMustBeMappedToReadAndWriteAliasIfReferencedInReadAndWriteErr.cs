
        public object F1(ref object a)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object a_vref = a;
            try
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1argp(this, _.NnO(_env.WScript, "WScript"), "Echo", _.ARGS.Ref(a_vref, v => { a_vref = v; }));
                });
            }
            finally { a = a_vref; }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
