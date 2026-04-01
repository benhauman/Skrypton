
            var enumerationContent = _.ENUMERABLE(_env.values).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                _env.value = enumerationContent.Current;
                _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.ARGS.Ref(_env.value, v => { _env.value = v; }));
            }
