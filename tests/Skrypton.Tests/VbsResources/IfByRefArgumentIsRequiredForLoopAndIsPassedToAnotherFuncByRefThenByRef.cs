
        public object F1(ref object x)
        {
            object F1_retVal = null;
            object i = null;
            object loopEnd = 0, loopStart = 0;
            var loopConstraintsInitialized = false;
            object x_vref = x;
            try
            {
                    loopEnd = _.NUM(_.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(x_vref, v => { x_vref = v; })));
                    loopStart = _.NUM((Int16)1);
                    if ((loopStart is DateTime) || (loopStart is Decimal))
                        i = loopStart;
                    loopStart = _.NUM((Int16)1, loopEnd);
                    loopConstraintsInitialized = true;
            }
            finally { x = x_vref; }
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
                {
                }
            }
            return F1_retVal;
        }
        public object F2(ref object value)
        {
            return _.VAL(value);
        }
