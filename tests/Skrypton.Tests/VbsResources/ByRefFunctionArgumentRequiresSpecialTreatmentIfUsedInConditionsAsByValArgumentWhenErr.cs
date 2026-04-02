
        public object F1(ref object a)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            bool ifResult;
            object a_zref = a;
            ifResult = _.IF(() => _.CALLm1v1(this, _outer, "F2", _.CALLm1v0(this, a_zref ?? throw new InvalidOperationException("Reference not set:a_zref"), "Name")), errOn);
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
