
        public object F1(ref object x)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object x_vref = x;
            try
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", _.TYPENAME(x_vref));
                });
            }
            finally { x = x_vref; }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
