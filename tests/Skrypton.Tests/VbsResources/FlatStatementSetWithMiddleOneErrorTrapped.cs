            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Test1");
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Test2");
            });
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Test3");
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
