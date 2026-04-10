
            if (_.IF(_.EQ(_env.x, "(")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Open");
            }
            else if (_.IF(_.EQ(_env.x, ")")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Close");
            }
            else if (_.IF(_.EQ(_env.x, ",")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, "WScript"), "Echo", "Split");
            }
