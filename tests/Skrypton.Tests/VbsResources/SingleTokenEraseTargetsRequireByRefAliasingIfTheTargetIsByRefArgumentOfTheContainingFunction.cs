
        public object F1(ref object a)
        {
            object F1_retVal = null;
            object byrefalias = a;
            try
            {
                _.ERASE(byrefalias, v => { byrefalias = v; });
            }
            finally { a = byrefalias; }
            return F1_retVal;
        }
