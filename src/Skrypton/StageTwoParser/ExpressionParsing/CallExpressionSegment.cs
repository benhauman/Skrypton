using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    /// <summary>
    /// A standalone CallExpressionSegment is essentially a specialised version of the CallSetItemExpressionSegment where there must be at least one Member
    /// Access Token.
    /// </summary>
    public sealed class CallExpressionSegment : CallSetItemExpressionSegment
    {
        public CallExpressionSegment(IReadOnlyCollection<IToken> memberAccessTokens, IReadOnlyCollection<ParsingExpression> arguments, ArgumentBracketPresenceOptions? zeroArgumentBracketsPresence)
            : base(memberAccessTokens, arguments, zeroArgumentBracketsPresence)
        {
            if (base.MemberAccessTokens.Count == 0)
                throw new ArgumentException("The memberAccessTokens set may not be empty");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references. There should be considered to be implicit MemberAccessorPointTokens between each
        /// token here (this will never contain any MemberAccessorOrDecimalPointToken references). The only token types that may be present in this data are
        /// BuiltInFunctionToken, BuiltInValueToken, KeyWordToken and NameToken.
        /// </summary>
        public new IReadOnlyCollection<IToken> MemberAccessTokens { get { return base.MemberAccessTokens; } }
    }
}
