
        public object F1(ref object a)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            bool ifResult;
            object a_vref = a;
            try
            {
                ifResult = _.IF(() => _.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(a_vref, v2 => { a_vref = v2; })), errOn);
            }
            finally { a = a_vref; }
            if (ifResult)
            {
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
        public object F2(object a)
        {
            return null;
        }
