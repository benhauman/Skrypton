            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            while (_.IF(() => _.IF(_.DIV((Int16)1, (Int16)0)), errOn))
            {
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
