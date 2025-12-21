using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.CodeBlocks;

namespace Skrypton.Tests.LegacyParser.Helpers
{
    public class CodeBlockSetComparer : IEqualityComparer<IEnumerable<ICodeBlock>>
    {
        public bool Equals(IEnumerable<ICodeBlock> x, IEnumerable<ICodeBlock> y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var arrayX = x.ToArray();
            var arrayY = y.ToArray();
            if (arrayX.Length != arrayY.Length)
                return false;

            var comparer = new CodeBlockComparer();
            for (var index = 0; index < arrayX.Length; index++)
            {
                if (!comparer.Equals(arrayX[index], arrayY[index]))
                    return false;
            }
            return true;
        }

        public int GetHashCode(IEnumerable<ICodeBlock> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
