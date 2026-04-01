
            if (_.IF(_.EQ(_env.x, "(")))
            {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Open");
            }
            else if (_.IF(_.EQ(_env.x, ")")))
            {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Close");
            }
            else if (_.IF(_.EQ(_env.x, ",")))
            {
                _.CALLm1v1(this, _env.WScript ?? throw new InvalidOperationException("Reference not set:"), "Echo", "Split");
            }
