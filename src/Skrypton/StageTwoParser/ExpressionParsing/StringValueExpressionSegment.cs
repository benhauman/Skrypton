using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class StringValueExpressionSegment : IExpressionSegment
    {
        public StringValueExpressionSegment(StringToken token)
        {
            Token = token ?? throw new ArgumentNullException(nameof(token));
        }
        public int LineIndex => Token.LineIndex;
        /// <summary>
        /// This will never be null
        /// </summary>
        public StringToken Token { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
#pragma warning disable CA1033 // Interface methods should be callable by child types
        IEnumerable<IToken> IExpressionSegment.AllTokens { get { return new[] { Token }; } }
#pragma warning restore CA1033 // Interface methods should be callable by child types

        public string RenderedContent
        {
            get { return "\"" + Token.Content + "\""; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
