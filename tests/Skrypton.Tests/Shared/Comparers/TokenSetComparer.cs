using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.Tests.Shared.Comparers
{
    public class TokenSetComparer : IEqualityComparer<IEnumerable<IToken>>
    {
        internal static readonly TokenSetComparer Instance = new TokenSetComparer();
        public TokenSetComparer() // @lubo: make it public
        {

        }

        public bool Equals(IEnumerable<IToken> x, IEnumerable<IToken> y)
        {
            if (x == null)
                throw new ArgumentNullException("x");
            if (y == null)
                throw new ArgumentNullException("y");

            var tokensArrayX = x.ToArray();
            var tokensArrayY = y.ToArray();
            if (tokensArrayX.Length != tokensArrayY.Length)
                return false;

            var tokenComparer = TokenComparer.Instance;
            for (var index = 0; index < tokensArrayX.Length; index++)
            {
                var token_X = tokensArrayX[index];
                var token_Y = tokensArrayY[index];
                if (!tokenComparer.Equals(token_X, token_Y))
                    return false;
            }
            return true;
        }

        public int GetHashCode(IEnumerable<IToken> obj)
        {
            if (obj == null)
                throw new ArgumentNullException("obj");

            return 0;
        }
    }
}
