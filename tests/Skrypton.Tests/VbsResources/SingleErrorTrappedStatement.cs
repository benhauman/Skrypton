            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Test1");
            });
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
