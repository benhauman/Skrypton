using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public interface IExpressionSegment
    {
        string RenderedContent { get; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        IEnumerable<IToken> AllTokens { get; }
    }
}
