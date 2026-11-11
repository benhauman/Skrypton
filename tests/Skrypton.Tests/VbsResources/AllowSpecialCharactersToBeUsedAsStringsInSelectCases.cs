
            if (_.IF(_.EQ(_env.x, "(")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Open");
            }
            else if (_.IF(_.EQ(_env.x, ")")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Close");
            }
            else if (_.IF(_.EQ(_env.x, ",")))
            {
                _.CALLm1v1(this, _.NnO(_env.WScript, null), "Echo", "Split");
            }
