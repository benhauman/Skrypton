            int errOn = _.GETERRORTRAPPINGTOKEN();

            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            IEnumerator enumerationContent = null;
            _.HANDLEERROR(errOn, () => {
                enumerationContent = _.ENUMERABLE(_env.values).GetEnumerator();
            });
            while (true)
            {
                if (enumerationContent != null)
                {
                    if (!enumerationContent.MoveNext())
                        break;
                    _env.value = enumerationContent.Current;
                }
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.ARGS.Ref(_env.value, v => { _env.value = v; }));
                });
                if (enumerationContent == null)
                    break;
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
