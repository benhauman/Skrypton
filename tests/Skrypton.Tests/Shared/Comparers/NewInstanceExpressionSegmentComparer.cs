using System;
using System.Collections.Generic;
using Skrypton.StageTwoParser.ExpressionParsing;

namespace Skrypton.Tests.Shared.Comparers
{
    public class NewInstanceExpressionSegmentComparer : IEqualityComparer<NewInstanceExpressionSegment>
    {
        public bool Equals(NewInstanceExpressionSegment x, NewInstanceExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            return x.ClassName.Content.Equals(y.ClassName.Content, StringComparison.InvariantCultureIgnoreCase);
        }

        public int GetHashCode(NewInstanceExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            return 0;
        }
    }
}
