            int errOn = _.GETERRORTRAPPINGTOKEN();

            if (_.IF(false))
            {
                _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            }
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Test1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
