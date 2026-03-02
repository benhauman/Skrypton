using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Handlers
{
    internal sealed class SelectHandler : AbstractBlockHandler
    {
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock Process(List<IToken> tokens)
        {
            if (tokens == null)
            {
                throw new ArgumentNullException(nameof(tokens));
            }

            if (tokens.Count == 0)
            {
                return null;
            }

            if (!checkAtomTokenPattern(tokens, ["SELECT", "CASE"], false))
            {
                return null;
            }

            // Trim out "SELECT CASE" tokens
            tokens.RemoveRange(0, 2);

            // Grab content for the case expression
            List<IToken> expressionTokens = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                if (isEndOfStatement(tokens, index))
                {
                    // Remove expression tokens (plus end-of-statement) from stream
                    tokens.RemoveRange(0, expressionTokens.Count + 1);
                    break;
                }

                if (index == 1)
                {
                    throw new NotImplementedException($"Multiple tokens on the 'expression' line. Line:{tokens.ElementAt(index).LineIndex}");
                }

                // Add token to expression (must be Atom or String)
                expressionTokens.Add(getTokenAtomOrDateStringLiteralOnly(tokens, index));
            }

            // Look for the first CASE entry (note: it's allowable for there to be no
            // CASE entries at all, and case entries can be empty). It is also valid
            // to have comments outside of the CASE entries, though no other tokens
            // are valid in those areas.
            List<CommentStatement> openingComments = new List<CommentStatement>();
            List<IToken> tokensIgnored = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                IToken token = tokens[index];
                if (token is CommentToken)
                {
                    openingComments.Add(new CommentStatement(token.Content, token.LineIndex));
                }
                else if (token is AbstractEndOfStatementToken)
                {
                    // Ignore blank lines
                    tokensIgnored.Add(token);
                }
                else if (token is AtomToken)
                {
                    if (token.Content.Equals("CASE", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                    else if (token.Content.Equals("END", StringComparison.OrdinalIgnoreCase))
                    {
                        if (index == (tokens.Count - 1))
                        {
                            throw new InvalidOperationException("Error processing SELECT CASE block - reached end of token stream");
                        }

                        IToken tokenNext = tokens[index + 1];
                        if (!(tokenNext is AtomToken))
                        {
                            throw new InvalidOperationException("Error processing SELECT CASE block - reached END followed invalid token [" + tokenNext.GetType().ToString() + "]");
                        }

                        if (!tokenNext.Content.Equals("SELECT", StringComparison.OrdinalIgnoreCase))
                        {
                            throw new InvalidOperationException("Error processing SELECT CASE block - reached non-SELECT END tokens");
                        }

                        break;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Invalid token encountered in SELECT CASE block [" + token.GetType().ToString() + "]");
                }
            }
            tokens.RemoveRange(0, openingComments.Count + tokensIgnored.Count);

            // Unless we hit "END SELECT" straight away, process CASE blocks
            List<SelectBlock.CaseBlockSegment> content = new List<SelectBlock.CaseBlockSegment>();
            if (!tokens[0].Content.Equals("END", StringComparison.OrdinalIgnoreCase))
            {
                string[]? endSequenceMet;
                CodeBlockHandler codeBlockHandler = new CodeBlockHandler([["CASE"], ["END", "SELECT"]]);
                while (true)
                {
                    // Try to grab value(s) for CASE block
                    // - Get lists of tokens (may be multiple values, may be ELSE..)
                    List<List<IToken>> exprValues = getEntryList(tokens, 1, new EndOfStatementNewLineToken(tokens[0].LineIndex));

                    // - Remove the CASE token
                    tokens.RemoveRange(0, 1);
                    // - Remove the exprValues tokens
                    bool doRemoveEofToken = true;
                    foreach (List<IToken> valueTokens in exprValues)
                    {
                        if (valueTokens.Count > 1)
                        {
                            // VBScript uses new lines to determine where one statement starts and another one begins, but you can use a colon to terminate a statement instead which allows you to span multiple statements across one line.
                            //
                            //     For example:
                            //
                            //          Case 0: flag = "af": country = "Afghanistan"
                            //     Is the equivalent of:
                            //
                            //          Case 0
                            //            flag = "af"
                            //            country = "Afghanistan"
                            // + https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/select-case-statement
                            // You can use multiple expressions or ranges in each Case clause. For example, the following line is valid:
                            //   Case 1 To 4, 7 To 9, 11, 13, Is > maxNumber
                            ////////////////////////////////////////////
                            var firstToken = valueTokens[0]; // StringToken - not 'Is' : The Is keyword used in the Case and Case Else statements is not the same as the Is Operator, which is used for object reference comparison.
                            var secondToken = valueTokens[1]; // NameToken : NOT a conditional operator after 'Is' or comma ','

                            if (firstToken is StringToken strtkn && secondToken is NameToken nametkn)
                            {
                                // Test with 'SelectCaseWithStringTokens'
                                // condition token and statement tokens on the same line
                                //    => remove only the condition token and all other token interpret as statement tokens.
                                doRemoveEofToken = false;
                                valueTokens.RemoveRange(1, valueTokens.Count - 1);
                                tokens.RemoveRange(0, valueTokens.Count);
                            }
                            else
                            {
                                throw new NotImplementedException($"Multiple tokens on the 'case' line. Line:{valueTokens[0].LineIndex}");
                            }
                        }
                        else
                        {
                            tokens.RemoveRange(0, valueTokens.Count);
                        }
                    }

                    // - Remove the commas between expressions
                    tokens.RemoveRange(0, exprValues.Count - 1);
                    if (doRemoveEofToken)
                    {
                        // - Remove the end-of-statement token
                        tokens.RemoveRange(0, 1);
                    }

                    // Quick check that it appears valid
                    bool caseElse = false;
                    if (exprValues.Count == 0)
                    {
                        throw new InvalidOperationException("CASE block with no comparison value");
                    }
                    else
                    {
                        IToken firstExprToken = exprValues[0][0];
                        if ((firstExprToken is AtomToken) && (firstExprToken.Content.Equals("ELSE", StringComparison.OrdinalIgnoreCase)))
                        {
                            if ((exprValues.Count > 1) || (exprValues[0].Count != 1))
                            {
                                throw new InvalidOperationException("Invalid CASE ELSE opening statement");
                            }

                            caseElse = true;
                        }
                    }

                    // Try to grab single CASE block content
                    List<ICodeBlock> blockContent = codeBlockHandler.Process(tokens, out endSequenceMet);
                    if (endSequenceMet == null)
                    {
                        throw new InvalidOperationException("Didn't find end sequence!");
                    }

                    // Add to CASE block list
                    if (caseElse)
                    {
                        content.Add(new SelectBlock.CaseBlockElseSegment(blockContent));
                    }
                    else
                    {
                        List<Expression> values = new List<Expression>();
                        foreach (List<IToken> valueTokens in exprValues)
                        {
                            values.Add(new Expression(valueTokens));
                        }

                        content.Add(new SelectBlock.CaseBlockExpressionSegment(values, blockContent));
                    }

                    // If we hit END SELECT then break out of loop, otherwise
                    // go back round to get the next block
                    if (endSequenceMet.Length == 2)
                    {
                        tokens.RemoveRange(0, endSequenceMet.Length);
                        if (tokens.Count > 0)
                        {
                            if (!(tokens[0] is AbstractEndOfStatementToken))
                            {
                                throw new InvalidOperationException("EndOfStatementToken missing after END FUNCTION");
                            }
                            else
                            {
                                tokens.RemoveAt(0);
                            }
                        }
                        break;
                    }

                }
            }

            // All done!
            return new SelectBlock(new Expression(expressionTokens), openingComments, content);
        }
    }
}
