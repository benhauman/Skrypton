
        public object F1(object x)
        {
            object F1_retVal = null;
            object i = null;
            var loopEnd = _.NUM(_.CALLm1argp(this, _outer ?? throw new InvalidOperationException("Reference not set:_outer"), "F2", _.ARGS.Ref(x, v => { x = v; })));
            var loopStart = _.NUM((Int16)1, loopEnd);
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
