
        public object F1(ref object x)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            object i = null;
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object loopEnd = 0, loopStart = 0;
            var loopConstraintsInitialized = false;
            object x_zref = x;
            _.HANDLEERROR(errOn, () => {
                loopEnd = _.NUM(_.ADD(x_zref, (Int16)1));
                loopStart = _.NUM((Int16)1);
                if ((loopStart is DateTime) || (loopStart is Decimal))
                    i = loopStart;
                loopStart = _.NUM((Int16)1, loopEnd);
                loopConstraintsInitialized = true;
            });
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
