using System;
using System.Collections.Generic;
using Skrypton.StageTwoParser.ExpressionParsing;

namespace Skrypton.Tests.Shared.Comparers
{
    public class OperatorOrComparisonExpressionSegmentComparer : IEqualityComparer<OperationExpressionSegment>
    {
        public bool Equals(OperationExpressionSegment x, OperationExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            return TokenComparer.Instance.Equals(x.Token, y.Token);
        }

        public int GetHashCode(OperationExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            return 0;
        }
    }
}
