
        public object F1(ref object a)
        {
            object F1_retVal = null;
            _.CALLm1v1(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F2", _.CONCAT("", a));
            return F1_retVal;
        }
        public object F2(ref object a)
        {
            return null;
        }
