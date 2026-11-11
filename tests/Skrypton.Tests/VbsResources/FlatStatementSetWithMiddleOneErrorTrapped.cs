            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Test1");
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Test2");
            });
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Test3");
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
