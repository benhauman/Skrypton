            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _env.i = (Int16)1;
            while (true)
            {
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.ARGS.Ref(_env.i, v => { _env.i = v; }));
                });
                var continueLoop = false;
                _.HANDLEERROR(errOn, () => {
                    _env.i = _.ADD(_env.i, (Int16)1);
                    continueLoop = _.StrictLTE(_env.i, 10);
                });
                if (!continueLoop)
                    break;
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
