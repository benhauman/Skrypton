
        public object F1(ref object x)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            object x_vref = x;
            try
            {
                x_vref = VBScriptConstants.Nothing;
            }
            finally { x = x_vref; }
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
