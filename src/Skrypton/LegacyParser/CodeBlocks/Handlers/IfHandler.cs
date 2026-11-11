using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.LegacyParser.CodeBlocks.Handlers
{
    /// <summary>
    /// This was one of the most complicated block handlers to implement! It could probably be tidied up some, but would need careful attention (and unit tests)
    /// </summary>
    internal sealed class IfHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock? Process(List<IToken> tokens)
        {
            // Input validation
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (tokens.Count == 0)
                return null;

            // Can we handle this point in the token stream?
            IToken firstToken = tokens[0];
            if ((!(firstToken is AtomToken)) || (!firstToken.Content.Equals("IF", StringComparison.OrdinalIgnoreCase)))
                return null;

            // Find the matching "THEN" for the "IF"
            List<IToken> expression = new List<IToken>();
            int indexThen = -1;
            for (int index = 0; index < tokens.Count; index++)
            {
                IToken token = tokens[index];
                if ((token is AtomToken) && (token.Content.Equals("THEN", StringComparison.OrdinalIgnoreCase)))
                {
                    indexThen = index;
                    break;
                }
                else
                    expression.Add(token);
            }
            if (indexThen == -1)
                throw new InvalidOperationException("IF statement not got matching THEN");

            // Is next token an end-of-statement token or is it a one-line IF statement?
            if (indexThen == (tokens.Count - 1))
                throw new InvalidOperationException("IF statement's THEN keyword is final token - missing content");
            if (tokens[indexThen + 1] is AbstractEndOfStatementToken)
                return ProcessMultiLine(tokens, indexThen);
            return ProcessSingleLine(tokens);
        }

        private static IfBlock ProcessMultiLine(List<IToken> tokens, int offsetToTHEN)
        {
            // ======================================================================
            // Notes on multi-line If's:
            // ======================================================================
            // - Invalid: Combining first statement with IF line
            //   eg. IF (1) THEN Func
            //        Func
            //       END IF
            //
            // - Invalid: Trailing same-line-end-of-statement after IF
            //   eg. IF (1) THEN :
            //        Func(0)
            //       END IF
            //
            // - Valid: Combining first statement with ELSEIF / ELSE line
            //   eg. IF (1) THEN
            //        Func
            //       ELSEIF (1) THEN Func
            //       ELSEIF (1) THEN
            //        Func
            //       ELSEIF (1) THEN Func
            //        Func
            //       ELSE Func
            //       END IF
            //
            // - Valid: Trailing same-line-end-of-statement after ELSEIF / ELSE
            //   eg. IF (1) THEN
            //        Func(0)
            //       ELSEIF(1) THEN :
            //        Func(0)
            //       END IF
            //
            // - Invalid: ELSE / ELSEIF must always be the first statement on the line
            //   eg. IF (0) THEN
            //        Func(1):Func(2):ELSE
            //       END IF
            //
            // - Valid: IF may be combined with a previous line
            //   eg. Func: IF (0) THEN
            //        Func:Func:ELSE
            //       END IF
            // ======================================================================
            // Grab content inside IF blocks
            string[][] endSequences = [["END", "IF"], ["ELSEIF"], ["ELSE"]];
            string[]? endSequenceMet = ["IF"];
            List<IfBlock.IfBlockSegment> ifContent = new List<IfBlock.IfBlockSegment>();
            List<IToken>? conditionTokens = null;
            while (true)
            {
                // First loop, we'll match this as the initial "IF" statement so we can
                // grab the content after it..
                if (endSequenceMet == null)
                    throw new InvalidOperationException("Didn't find end sequence!");

                // Remove the last end sequence tokens (on the first loop, this will
                // actually remove the initial "IF" token)
                tokens.RemoveRange(0, endSequenceMet.Length);

                // The sequences that indicate further content are single-token ("IF",
                // "ELSEIF", ELSE")
                if (endSequenceMet.Length == 1)
                {
                    // An "ELSE" token won't have any condition content
                    if (endSequenceMet[0].Equals("ELSE", StringComparison.OrdinalIgnoreCase))
                        conditionTokens = null;
                    else
                    {
                        // Grab the condition content (up to the "THEN" token) - this will
                        // throw an exception if it can't find "THEN" or there is no content
                        conditionTokens = getConditionContent(tokens).ToList();
                    }

                    // If the first character is a new-line end-of-statement (which is a
                    // common form), we may as well pull that out as well
                    if ((tokens.Count > 0) && (tokens[0] is AbstractEndOfStatementToken))
                        tokens.RemoveAt(0);
                }

                // Try to find the end of this content block (this will remove tokens
                // from the stream that have been processed by the block handler)
                var codeBlockHandler = new CodeBlockHandler(endSequences);
                var blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
                if (endSequenceMet == null)
                    throw new InvalidOperationException("Didn't find end sequence!");

                // Record this content block
                if (conditionTokens == null)
                    ifContent.Add(new IfBlock.IfBlockElseSegment(blockContent));
                else
                {
                    CodeExpression conditionStatement = new CodeExpression(conditionTokens);
                    ifContent.Add(new IfBlock.IfBlockConditionSegment(conditionStatement, blockContent));
                }

                // If we hit, "END IF" then we're all done
                if ((endSequenceMet.Length == 2) && (endSequenceMet[0].Equals("END", StringComparison.OrdinalIgnoreCase)))
                {
                    // Need to remove the final tokens - if we don't break out of the loop
                    // here, then end sequence tokens are removed up at the top
                    tokens.RemoveRange(0, endSequenceMet.Length);
                    break;
                }
            }

            // Expect the next token will be end-of-statement for IF block - drop that too
            if ((tokens.Count > 0) && (tokens[0] is AbstractEndOfStatementToken))
                tokens.RemoveAt(0);

            // Return fully-formed IfBlock
            return new IfBlock(ifContent);
        }

        private static IfBlock ProcessSingleLine(List<IToken> tokens)
        {
            // ======================================================================
            // Notes on single-line If's:
            // ======================================================================
            // - Invalid: Combining single line with multi-line
            //   eg. IF (1) THEN Func() ELSE IF (1) THEN
            //        Func()
            //       END IF
            //
            // - Invalid: ElseIf formations
            //   eg. IF (1) THEN Func() ELSEIF (2) THEN Func()
            //
            // - Combined If statements ARE valid
            //   eg. IF (1) THEN Func() ELSE IF (2) THEN Func()
            //   eg. IF (1) THEN IF (2) THEN Func()
            //
            // - Empty initial statement on If, with same-line end-of-statement
            //   eg. IF (1) THEN : Func(2)
            //
            // - Dropping statement from else is valid
            //   eg. IF (1) THEN Func() ELSE
            //
            // - Nested if statements can only ELSE to the last statement
            //   Valid:   IF (1) THEN Func() ELSE IF (1) THEN Func() ELSE Func()
            //   InValid: IF (1) THEN Func() ELSE IF (1) THEN Func() ELSE Func() ELSE Func()
            // ======================================================================
            // Grab content up to new line token (or end of content)
            List<IToken> ifTokens = new List<IToken>();
            foreach (IToken token in tokens)
            {
                if (token is EndOfStatementNewLineToken)
                    break;
                else
                {
                    if ((!(token is AtomToken))
                    && (!(token is DateLiteralToken))
                    && (!(token is StringToken))
                    && (!(token is CommentToken)) // also a comment after a string literal: If c Then r = "q" ' comment
                    && (!(token is EndOfStatementSameLineToken)))
                        throw new InvalidOperationException("IfHandler.ProcessSingleLine: Encountered invalid Token - should all be AtomToken, DateLiteralToken, StringToken or EndOfStatementSameLineToken until new-line end of statement");
                    ifTokens.Add(token);
                }
            }

            // - "END IF" may end a single-line IF (as Windows VBScript allows - issue #97):
            //   eg. IF (1) THEN Func() ELSE Func() END IF
            //   eg. IF (1) THEN Func() : END IF : Func()
            //   eg. FUNCTION F() : IF (1) THEN F = 1 END IF : END FUNCTION
            //   the statements after it (on the same line) are not part of the IF; an END IF
            //   ends the innermost IF ("IF (1) THEN F() ELSE IF (2) THEN G() END IF : H()" -
            //   H() is in the ELSE of the first IF)
            int offsetEndIf = -1;
            int nestedIfs = 0;
            for (int index = 1; index < ifTokens.Count; index++)
            {
                if (!(ifTokens[index] is AtomToken) || !ifTokens[index].Content.Equals("IF", StringComparison.OrdinalIgnoreCase))
                    continue;
                bool isEndIf = (ifTokens[index - 1] is AtomToken) && ifTokens[index - 1].Content.Equals("END", StringComparison.OrdinalIgnoreCase);
                if (!isEndIf)
                    nestedIfs++;
                else if (nestedIfs > 0)
                    nestedIfs--;
                else
                {
                    offsetEndIf = index - 1;
                    break;
                }
            }
            bool endsWithEndIf = offsetEndIf != -1;
            if (endsWithEndIf)
            {
                // only the end of the statement may follow END IF
                int afterEndIf = offsetEndIf + 2;
                if ((afterEndIf < ifTokens.Count) && !(ifTokens[afterEndIf] is AbstractEndOfStatementToken) && !(ifTokens[afterEndIf] is CommentToken))
                    throw new InvalidOperationException("Expected end of statement after END IF in single-line IF");
                ifTokens.RemoveRange(offsetEndIf, ifTokens.Count - offsetEndIf);
            }

            // We'll need to store the number of tokens associated with this IF statement
            // because we'll pull them out of the stream later (but we'll be manipulating
            // this ifTokens list as well, so stash the original count now) - with END IF,
            // the IF statement includes those two tokens
            int ifTokensCount = ifTokens.Count + (endsWithEndIf ? 2 : 0);

            // Pull "IF" token from the start
            ifTokens.RemoveAt(0);

            // Look for "THEN" token (pull handled tokens out of stream) - this will
            // throw an exception if it can't find "THEN" or there is no content
            var conditionTokens = getConditionContent(ifTokens).ToList();

            // Look for "ELSE" token (if present), ensuring we don't encounter any
            // "ELSEIF" tokens (which aren't valid in a single line "IF")
            int offsetElse = -1;
            for (int index = 0; index < ifTokens.Count; index++)
            {
                IToken token = ifTokens[index];
                if (token is AtomToken)
                {
                    if (token.Content.Equals("ELSE", StringComparison.OrdinalIgnoreCase))
                    {
                        offsetElse = index;
                        break;
                    }
                    else if (token.Content.Equals("ELSEIF", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Invalid content: ELSEIF found in single-line IF");
                }
            }

            // We have the condition statement, now get the Post-THEN content and
            // Post-ELSE content (may be null) - ie. the condition's "met" and
            // "not met" code blocks
            List<IToken> truthTokens;
            List<IToken>? notTokens;
            if (offsetElse == -1)
            {
                truthTokens = getTokenListSection(ifTokens, 0);
                notTokens = null;
            }
            else
            {
                truthTokens = getTokenListSection(ifTokens, 0, offsetElse);
                notTokens = getTokenListSection(ifTokens, offsetElse + 1);
            }

            // A ":" before END IF ends the last statement ("IF (1) THEN Func() : END IF")
            if (endsWithEndIf)
                TrimTrailingSameLineEndOfStatement(notTokens ?? truthTokens);

            // Note: It's not valid for Post-THEN content to be empty
            if (truthTokens.Count == 0)
                throw new InvalidOperationException("Empty THEN content in IF");

            // If we've got this far, we can pull out the processed tokens from the input list
            // - If statement ended at new-line-end-of-statement (as opposed to end of
            //   token stream) then remove that token as well; after END IF, the end of
            //   statement (":" or new line) that follows it
            tokens.RemoveRange(0, ifTokensCount);
            if ((tokens.Count > 0) && (!endsWithEndIf || (tokens[0] is AbstractEndOfStatementToken)))
                tokens.RemoveAt(0);

            // Translate token sets into code blocks
            CodeBlockHandler codeBlockHandler = CodeBlockHandler.RootBlock;
            List<ICodeBlock> truthStatement = codeBlockHandler.Process(truthTokens, out var _);
            List<ICodeBlock>? notStatement;
            if (notTokens == null)
            {
                notStatement = null;
            }
            else
            {
                notStatement = codeBlockHandler.Process(notTokens, out var _);
            }

            CodeExpression conditionStatement = new CodeExpression(conditionTokens);

            // Generate IfBlock
            List<IfBlock.IfBlockSegment> ifContent = new List<IfBlock.IfBlockSegment>();
            ifContent.Add(new IfBlock.IfBlockConditionSegment(conditionStatement, truthStatement));
            if (notStatement != null)
                ifContent.Add(new IfBlock.IfBlockElseSegment(notStatement));
            return new IfBlock(ifContent);
        }

        private static void TrimTrailingSameLineEndOfStatement(List<IToken> tokens)
        {
            while ((tokens.Count > 0) && (tokens[tokens.Count - 1] is EndOfStatementSameLineToken))
                tokens.RemoveAt(tokens.Count - 1);
        }

        /// <summary>
        /// While processing an IF / ELSEIF statement - that keyword should have been removed
        /// from the token stream, such that the start of the stream is the start of the
        /// condition content. This will trim off and return the condition content. An
        /// exception will be thrown for invalid content, but neither null nor an
        /// empty list will ever be returned.
        /// </summary>
        private static IToken[] getConditionContent(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));

            // Look for "THEN" token - only acceptable tokens are Atom or literal (String or Date)
            int offsetThen = -1;
            for (var index = 0; index < tokens.Count; index++)
            {
                var token = tokens[index];
                if ((!(token is AtomToken)) && (!(token is DateLiteralToken)) && (!(token is StringToken)))
                    throw new InvalidOperationException("Encountered invalid token looking for THEN content");
                if ((token is AtomToken) && (token.Content.Equals("THEN", StringComparison.OrdinalIgnoreCase)))
                {
                    offsetThen = index;
                    break;
                }
            }
            if (offsetThen == -1)
                throw new InvalidOperationException("Invalid content: no THEN token for IF / ELSEIF");

            // Grab condition content
            var conditionTokens = getTokenListSection(tokens, 0, offsetThen);

            // Trim out the handled tokens and the "THEN"
            tokens.RemoveRange(0, conditionTokens.Count + 1);

            // Return the content
            return conditionTokens.ToArray();
        }
    }
}
