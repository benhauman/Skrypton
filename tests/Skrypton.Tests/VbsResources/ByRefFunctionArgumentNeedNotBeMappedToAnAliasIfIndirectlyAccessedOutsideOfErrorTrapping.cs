
        public object F1(ref object a)
        {
            object F1_retVal = null;
            _.CALLm1v1(this, _outer, "F2", _.CALLm1v0(this, a ?? throw new InvalidOperationException("Reference not set:a"), "Name"));
            return F1_retVal;
        }
        public object F2(ref object a)
        {
            return null;
        }
