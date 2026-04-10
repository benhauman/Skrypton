
            var enumerationContent = _.ENUMERABLE(_env.values).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                _env.value = enumerationContent.Current;
                _.CALLm1argp(this, _.NnO(_env.WScript, "WScript"), "Echo", _.ARGS.Ref(_env.value, v => { _env.value = v; }));
            }
