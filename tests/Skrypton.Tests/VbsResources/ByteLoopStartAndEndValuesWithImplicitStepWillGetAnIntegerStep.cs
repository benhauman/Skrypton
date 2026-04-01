
            var loopEnd = _.CBYTE(5);
            var loopStart = _.NUM(_.CBYTE(1), loopEnd, (Int16)1);
            if (_.StrictLTE(loopStart, loopEnd))
            {
                for (_outer.i = loopStart; _.StrictLTE(_outer.i, loopEnd); _outer.i = _.ADD(_outer.i, (Int16)1))
                {
                }
            }
