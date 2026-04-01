            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Test1");
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Test2");
            });
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Test3");
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
