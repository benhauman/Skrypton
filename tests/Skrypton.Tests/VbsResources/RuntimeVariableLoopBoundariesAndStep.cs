
            var loopEnd = _.NUM(_env.b);
            var loopStep = _.NUM(_env.c);
            var loopStart = _.NUM(_env.a, loopEnd, loopStep);
            if ((_.StrictLTE(loopStart, loopEnd) && _.StrictGTE(loopStep, 0))
            || (_.StrictGT(loopStart, loopEnd) && _.StrictLT(loopStep, 0)))
            {
                for (_env.i = loopStart;
                    (_.StrictGTE(loopStep, 0) && _.StrictLTE(_env.i, loopEnd)) || (_.StrictLT(loopStep, 0) && _.StrictGTE(_env.i, loopEnd));
                    _env.i = _.ADD(_env.i, loopStep))
                {
                }
            }
