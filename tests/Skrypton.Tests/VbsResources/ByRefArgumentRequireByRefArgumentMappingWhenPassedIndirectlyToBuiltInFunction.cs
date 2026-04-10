
        public object F1(ref object x)
        {
            object F1_retVal = null;
            object x_vref = x;
            try
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", _.TYPENAME(_.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(x_vref, v => { x_vref = v; }))));
            }
            finally { x = x_vref; }
            return F1_retVal;
        }
        public object F2(ref object x)
        {
            return _.VAL(x);
        }
