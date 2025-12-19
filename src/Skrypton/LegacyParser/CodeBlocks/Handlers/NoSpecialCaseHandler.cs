using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Handlers
{
    public class NoSpecialCaseHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException("tokens");
            if (tokens.Count > 0)
            {
                IToken token = tokens[0];
                if (token is InlineCommentToken)
                {
                    tokens.RemoveAt(0);
                    return new InlineCommentStatement(token.Content, token.LineIndex);
                }
                if (token is CommentToken)
                {
                    tokens.RemoveAt(0);
                    return new CommentStatement(token.Content, token.LineIndex);
                }
                if (token is AbstractEndOfStatementToken)
                {
                    tokens.RemoveAt(0);
                    return new BlankLine();
                }
            }
            return null;
        }
    }
}
