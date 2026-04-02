
        public object F1(ref object x)
        {
            object F1_retVal = null;
            if (_.IF(true))
            {
            }
            else
            {
                bool ifResult;
                object x_vref = x;
                try
                {
                    ifResult = _.IF(_.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(x_vref, v2 => { x_vref = v2; })));
                }
                finally { x = x_vref; }
                if (ifResult)
                {
                }
            }
            return F1_retVal;
        }
        public object F2(ref object x)
        {
            return null;
        }
