using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Handlers
{
    internal sealed class ClassHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (tokens.Count < 3)
                return null;

            // Look for start of function declaration
            if (!checkAtomTokenPattern(tokens, "CLASS", false))
                return null;
            if (!(tokens[1] is AtomToken))
                return null;
            if (!(tokens[2] is AbstractEndOfStatementToken))
                return null;
            var classNameToken = tokens[1];
            tokens.RemoveRange(0, 3);

            // Get function content
            string[] endSequenceMet;
            var codeBlockHandler = new CodeBlockHandler(["END", "CLASS"]);
            var functionContent = codeBlockHandler.Process(tokens, out endSequenceMet);
            if (endSequenceMet == null)
                throw new InvalidOperationException("Didn't find encounter end sequence!");

            // Remove end sequence tokens
            tokens.RemoveRange(0, endSequenceMet.Length);
            if (tokens.Count > 0)
            {
                if (!(tokens[0] is AbstractEndOfStatementToken))
                    throw new InvalidOperationException("EndOfStatementToken missing after END CLASS");
                else
                    tokens.RemoveAt(0);
            }

            return new ClassBlock(new NameToken(false, classNameToken.ContentUpperX(), classNameToken.LineIndex), functionContent);
        }
    }
}
