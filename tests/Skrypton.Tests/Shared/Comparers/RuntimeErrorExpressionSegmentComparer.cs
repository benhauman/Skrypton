using System;
using System.Collections.Generic;
using Skrypton.StageTwoParser.ExpressionParsing;

namespace Skrypton.Tests.Shared.Comparers
{
    public class RuntimeErrorExpressionSegmentComparer : IEqualityComparer<RuntimeErrorExpressionSegment>
    {
        public bool Equals(RuntimeErrorExpressionSegment x, RuntimeErrorExpressionSegment y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            return
                (x.RenderedContent == y.RenderedContent) &&
                (x.ExceptionType == y.ExceptionType) &&
                (x.Message == y.Message) &&
                new TokenSetComparer().Equals(x.AllTokens, y.AllTokens);
        }

        public int GetHashCode(RuntimeErrorExpressionSegment obj)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            return 0;
        }
    }
}
