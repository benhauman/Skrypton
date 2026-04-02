
        public object F1(ref object a)
        {
            object F1_retVal = null;
            object a_vref3 = a;
            try
            {
                F1_retVal = _.VAL(_.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(a_vref3, v => { a_vref3 = v; })));
            }
            finally { a = a_vref3; }
            return F1_retVal;
        }
        public object F2(ref object a)
        {
            return null;
        }
