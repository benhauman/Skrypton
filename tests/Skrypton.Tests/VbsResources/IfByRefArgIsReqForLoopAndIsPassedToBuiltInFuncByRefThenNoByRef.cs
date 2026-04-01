
        public object F1(ref object x)
        {
            object F1_retVal = null;
            object i = null;
            var loopEnd = _.UBOUND(x);
            var loopStart = _.NUM((Int16)1, loopEnd);
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
                {
                }
            }
            return F1_retVal;
        }
