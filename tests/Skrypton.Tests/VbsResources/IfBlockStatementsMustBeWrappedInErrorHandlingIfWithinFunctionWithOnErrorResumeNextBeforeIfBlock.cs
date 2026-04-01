
        public object F1(object value)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            if (_.IF(() => true, errOn))
            {
                _.HANDLEERROR(errOn, () => {
                    F1_retVal = _.DATEVALUE(value);
                });
            }
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
