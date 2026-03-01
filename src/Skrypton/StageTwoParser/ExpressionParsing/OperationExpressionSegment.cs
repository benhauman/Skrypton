using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class OperationExpressionSegment : IExpressionSegment
    {
        public OperationExpressionSegment(OperatorToken token)
        {
            Token = token ?? throw new ArgumentNullException(nameof(token));
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public OperatorToken Token { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
#pragma warning disable CA1033 // Interface methods should be callable by child types
        IEnumerable<IToken> IExpressionSegment.AllTokens { get { return new[] { Token }; } }
#pragma warning restore CA1033 // Interface methods should be callable by child types

        public string RenderedContent
        {
            get { return Token.Content; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
