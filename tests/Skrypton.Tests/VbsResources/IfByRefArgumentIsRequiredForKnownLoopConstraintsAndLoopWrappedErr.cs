
        public object F1(ref object x)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            object i = null;
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object loopEnd = 0, loopStart = 0;
            var loopConstraintsInitialized = false;
            object x_vref = x;
            try
            {
                _.HANDLEERROR(errOn, () => {
                    loopEnd = _.NUM(_.CALLm1argp(this, _outer, "F2", _.ARGS.Ref(x_vref, v => { x_vref = v; })));
                    loopStart = _.NUM((Int16)1);
                    if ((loopStart is DateTime) || (loopStart is Decimal))
                        i = loopStart;
                    loopStart = _.NUM((Int16)1, loopEnd);
                    loopConstraintsInitialized = true;
                });
            }
            finally { x = x_vref; }
            if (_.StrictLTE(loopStart, loopEnd))
            {
                if (loopConstraintsInitialized)
                    i = loopStart;
                while (true)
                {
                    if (!loopConstraintsInitialized)
                        break;
                    var continueLoop = false;
                    _.HANDLEERROR(errOn, () => {
                        i = _.ADD(i, (Int16)1);
                        continueLoop = _.StrictLTE(i, loopEnd);
                    });
                    if (!continueLoop)
                        break;
                }
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
        public object F2(ref object value)
        {
            object F2_retVal = null;
            F2_retVal = _.VAL(value);
            value = (Int16)123;
            return F2_retVal;
        }
