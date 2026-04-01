            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            object loopEnd = 0, loopStart = 0;
            var loopConstraintsInitialized = false;
            _.HANDLEERROR(errOn, () => {
                loopEnd = _.NUM(_env.b);
                loopStart = _.NUM(_env.a);
                if ((loopStart is DateTime) || (loopStart is Decimal))
                    _env.i = loopStart;
                loopStart = _.NUM(_env.a, loopEnd, (Int16)1);
                loopConstraintsInitialized = true;
            });
            if (_.StrictLTE(loopStart, loopEnd))
            {
                if (loopConstraintsInitialized)
                    _env.i = loopStart;
                while (true)
                {
                    _.HANDLEERROR(errOn, () => {
                        _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", _.ARGS.Ref(_env.i, v => { _env.i = v; }));
                    });
                    if (!loopConstraintsInitialized)
                        break;
                    var continueLoop = false;
                    _.HANDLEERROR(errOn, () => {
                        _env.i = _.ADD(_env.i, (Int16)1);
                        continueLoop = _.StrictLTE(_env.i, loopEnd);
                    });
                    if (!continueLoop)
                        break;
                }
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
