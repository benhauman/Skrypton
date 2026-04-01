
        public object F1(object value)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            object i = null;
            i = (Int16)1;
            while (true)
            {
                if (_.IF(() => true, errOn))
                {
                    _.HANDLEERROR(errOn, () => {
                        F1_retVal = _.DATEVALUE(value);
                    });
                }
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
                var continueLoop = false;
                _.HANDLEERROR(errOn, () => {
                    i = _.ADD(i, (Int16)1);
                    continueLoop = _.StrictLTE(i, 1);
                });
                if (!continueLoop)
                    break;
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
