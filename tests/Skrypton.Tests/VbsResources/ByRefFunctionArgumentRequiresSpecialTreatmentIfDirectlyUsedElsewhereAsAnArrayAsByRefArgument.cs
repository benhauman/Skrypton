
        public object F1(ref object a)
        {
            object F1_retVal = null;
            object a_vref = a;
            try
            {
                _.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F2", _.ARGS.RefIfArray(a_vref, _.ARGS.Val((Int16)0)));
            }
            finally { a = a_vref; }
            return F1_retVal;
        }
        public object F2(ref object a)
        {
            return null;
        }
