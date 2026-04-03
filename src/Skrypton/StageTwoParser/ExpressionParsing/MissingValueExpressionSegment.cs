using System;
using System.Collections.Generic;
using System.Text;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    internal sealed class MissingValueExpressionSegment : IExpressionSegment
    {
        public MissingValueExpressionSegment(ArgumentSeparatorToken argumentSeparator)
        {
            LineIndex = argumentSeparator.LineIndex;
            AllTokens = Array.Empty<IToken>();
            RenderedContent = "miiiissing value";
        }
        public int LineIndex { get; }
        public string RenderedContent { get; }
        public IEnumerable<IToken> AllTokens { get; }
    }
}
