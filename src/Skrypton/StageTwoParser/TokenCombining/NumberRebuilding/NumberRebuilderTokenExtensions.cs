using System;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.TokenCombining.NumberRebuilding
{
    public static class NumberRebuilderTokenExtensions
    {
        public static bool Is<T>(this IToken token) where T : IToken
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return (token.GetType() == typeof(T));
        }

        public static bool IsMinusSignOperator(this IToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            return token.Is<OperatorToken>() && (token.Content == "-");
        }
    }
}
