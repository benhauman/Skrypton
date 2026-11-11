            int errOn = _.GETERRORTRAPPINGTOKEN();

            if (_.IF(false))
            {
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            }
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Test1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
