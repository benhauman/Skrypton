
        public object F1(object value)
        {
            object F1_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            if (_.IF(true))
            {
                F1_retVal = _.DATEVALUE(value);
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return F1_retVal;
            }
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                F1_retVal = _.DATE();
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return F1_retVal;
        }
