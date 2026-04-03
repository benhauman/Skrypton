using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    [DebuggerDisplay("{LineIndex},,")]
    internal sealed class MissingValueExpressionSegment : IExpressionSegment
    {
        public MissingValueExpressionSegment(ArgumentSeparatorToken argumentSeparator)
        {
            LineIndex = argumentSeparator.LineIndex;
        }
        public int LineIndex { get; }
        public string RenderedContent => "<miiiissing value";
        public IEnumerable<IToken> AllTokens => Array.Empty<IToken>();
    }
}
