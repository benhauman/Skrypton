using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.StageTwoParser.ExpressionParsing;

namespace Skrypton.Tests.Shared.Comparers
{
    public sealed class ExpressionSetComparer : IEqualityComparer<IEnumerable<ParsingExpression>>
    {
        public bool Equals(IEnumerable<ParsingExpression> x, IEnumerable<ParsingExpression> y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            var arrayX = x.ToArray();
            var arrayY = y.ToArray();
            if (arrayX.Length != arrayY.Length)
                return false;

            var expressionComparer = new ExpressionSegmentSetComparer();
            for (var index = 0; index < arrayX.Length; index++)
            {
                var xExpr = arrayX[index].Segments;
                var yExpr = arrayY[index].Segments;
                if (!expressionComparer.Equals(xExpr, yExpr))
                    return false;
            }
            return true;
        }

        public int GetHashCode(IEnumerable<ParsingExpression> obj)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            return 0;
        }
    }
}
