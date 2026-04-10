
        public object Render(ref object x)
        {
            object Render_retVal = null;
            var with = _.OBJ(x);
            _.CALLm1v1(this, _.NnO(with, "with"), "Draw", "Test");
            return Render_retVal;
        }
