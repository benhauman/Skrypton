
            var loopEnd = _.CBYTE(5);
            var loopStep = _.CBYTE(1);
            var loopStart = _.NUM(_.CBYTE(1), loopEnd, loopStep);
            if ((_.StrictLTE(loopStart, loopEnd) && _.StrictGTE(loopStep, 0))
            || (_.StrictGT(loopStart, loopEnd) && _.StrictLT(loopStep, 0)))
            {
                for (_outer.i = loopStart;
                    (_.StrictGTE(loopStep, 0) && _.StrictLTE(_outer.i, loopEnd)) || (_.StrictLT(loopStep, 0) && _.StrictGTE(_outer.i, loopEnd));
                    _outer.i = _.ADD(_outer.i, loopStep))
                {
                }
            }
