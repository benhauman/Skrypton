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
                    {
                        _env.value = null;
                        break;
                    }
                    _env.value = enumerationContent.Current;
                }
                _.HANDLEERROR(errOn, () => {
                    _.CALLm1argp(this, _.NnO(_env.WScript, null), "Echo", _.ARGS.Ref(_env.value, v => { _env.value = v; }));
                });
                if (enumerationContent == null)
                    break;
            }
            _.RELEASEERRORTRAPPINGTOKEN(errOn);
