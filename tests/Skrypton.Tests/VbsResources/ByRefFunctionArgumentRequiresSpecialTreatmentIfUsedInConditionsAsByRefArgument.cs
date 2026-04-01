
        public object F1(ref object a)
        {
            object F1_retVal = null;
            bool ifResult;
            object a_vref = a;
            try
            {
                ifResult = _.IF(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F2", _.ARGS.Ref(a_vref, v2 => { a_vref = v2; })));
            }
            finally { a = a_vref; }
            if (ifResult)
            {
            }
            return F1_retVal;
        }
        public object F2(ref object a)
        {
            return null;
        }
