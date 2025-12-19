using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class DateValueExpressionSegment : IExpressionSegment
    {
        public DateValueExpressionSegment(DateLiteralToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            Token = token;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public DateLiteralToken Token { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        IEnumerable<IToken> IExpressionSegment.AllTokens { get { return new[] { Token }; } }

        public string RenderedContent
        {
            get { return "#" + Token.Content + "#"; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
