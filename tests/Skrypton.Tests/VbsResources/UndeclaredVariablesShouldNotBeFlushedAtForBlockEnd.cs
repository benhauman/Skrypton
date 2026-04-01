
        public object F1()
        {
            object F1_retVal = null;
            object j = null; /* Undeclared in source */
            object i = null; /* Undeclared in source */
            for (i = (Int16)1; _.StrictLTE(i, 5); i = _.ADD(i, (Int16)1))
            {
                _.CALLm1argp(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.ARGS.Ref(j, v => { j = v; }));
            }
            return F1_retVal;
        }
