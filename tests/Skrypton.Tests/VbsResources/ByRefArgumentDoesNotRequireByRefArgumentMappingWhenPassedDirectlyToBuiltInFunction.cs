
        public object F1(ref object x)
        {
            object F1_retVal = null;
            _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:WScript"), "Echo", _.TYPENAME(x));
            return F1_retVal;
        }
