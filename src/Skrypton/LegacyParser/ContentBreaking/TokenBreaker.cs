using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.CodeBlocks.Handlers;

namespace Skrypton.LegacyParser.ContentBreaking
{
    public static class TokenBreaker
    {
        private const string TokenBreakChars = ",.*&+-/\\=!(){}[]<>:;\n";

        /// <summary>
        /// Break down an UnprocessedContentToken into a combination of AtomToken and AbstractEndOfStatementToken references. This will never return null nor a set
        /// containing any null references.
        /// </summary>
        public static IReadOnlyCollection<IToken> BreakUnprocessedToken(UnprocessedContentToken token)
        {
#pragma warning disable CA1820 // Test for empty strings using string length
            if (token == null)
                throw new ArgumentNullException(nameof(token));

            var lineIndex = token.LineIndex;
            var buffer = "";
            var content = token.Content;
            var tokens = new List<IToken>();
            bool? last_chr0IsWhitespace = null;
            for (var index = 0; index < content.Length; index++)
            {
                string debug_remaining_text = content.Substring(index);
                string chr = content.Substring(index, 1);
                bool chr0IsWhitespace = char.IsWhiteSpace(chr, 0);
                if (char.IsWhiteSpace(chr, 0) && (chr != "\n"))
                {
                    // If we've found a (non-line-return) whitespace character, push content retrieved from the token so far (if any), into a fresh token on the
                    // list and clear the buffer to accept following data.
                    if (buffer != "")
                    {
                        var tkn = AtomToken.GetNewToken(buffer.ToUpperX(), hasLeadingWhiteSpace: true, lineIndex);
                        var prevToken = tokens.LastOrDefault();
                        tokens.Add(tkn);
                        //if (prevToken is MemberAccessorOrDecimalPointToken && tkn is NameToken)
                        if (buffer == "") // false
                        {
                            tokens.Add(new WhiteSpaceToken(lineIndex)); //  inside of a WITH statement:  .MethodX .PropertyA
                        }
                    }
                    buffer = "";
                }
                else
                {
                    bool characterIsTokenBreaker;
                    if (TokenBreakChars.IndexOf(chr, StringComparison.Ordinal) != -1)
                    {
                        characterIsTokenBreaker = true;
                    }
                    else if (chr == "_")
                    {
                        // An underscore is a line return continuation character if it follows whitespace, but it must be part of a variable name if it is not
                        // preceded by whitespace (and line return continuation is a token-breaker, as opposed to an underscore that is part of the current
                        // token)
                        characterIsTokenBreaker = (index > 0) && char.IsWhiteSpace(content, index - 1);
                    }
                    else
                    {
                        characterIsTokenBreaker = false;
                    }

                    if (characterIsTokenBreaker)
                    {
                        // If the current character is a "&" then it may be a string concatenation or it may be the start of a hex number (eg. "&h001"), if it's
                        // the latter then we want to represent the content as a single token "&h001" not break the "&" out.
                        if ((chr == "&") && (index <= (content.Length - 3)))
                        {
                            var chrNext = content.Substring(index + 1, 1);
                            var chrNextNext = content.Substring(index + 2, 1);
                            if (chrNext.Equals("H", StringComparison.OrdinalIgnoreCase) && ("0123456789".IndexOf(chrNextNext, StringComparison.Ordinal) != -1))
                            {
                                buffer += chr;
                                continue;
                            }
                        }

                        // If we've found another "break" character (which means a token split is identified, but that we want to keep the break character itself,
                        // unlike with whitespace breaks), then do similar to above.
                        if (buffer != "")
                        {
                            bool canBeDimToken = DimHandler.CanBeHandledAsDimToken(tokens);
                            IToken newTkn = AtomToken.GetNewToken(buffer.ToUpperX(), last_chr0IsWhitespace ?? false, lineIndex);
                            if (canBeDimToken && newTkn is BuiltInFunctionToken binFun && binFun.FunctionId == BuiltInFunctionId.BuiltInFunctionSPACE) // VBScript:  "Dim Space : Space = 1" is valid, so "Space" in this context should be treated as a NameToken, not a BuiltInFunctionToken
                            {
                                newTkn = new NameToken(false, newTkn.ContentUpperX(), newTkn.LineIndex);
                            }
                            tokens.Add(newTkn);
                        }

                        bool hasLeadingWhiteSpace;
                        if (last_chr0IsWhitespace != null && last_chr0IsWhitespace.Value && !chr.ToUpperX().containsWhiteSpace() && AtomToken.isMemberAccessorUpper(chr.ToUpperX())) // isMemberAccessorUpper
                        {
                            IToken? lastTokenIfAny = tokens.LastOrDefault();
                            if (lastTokenIfAny == null)
                            {
                                hasLeadingWhiteSpace = false;
                            }
                            else
                            {
                                if (lastTokenIfAny is NameToken)
                                {
                                    hasLeadingWhiteSpace = true; // VBScript:
                                }
                                else if (lastTokenIfAny is KeyWordToken kw && kw.KeyWordId == KnownKeyWordId.KeywordSet)
                                {
                                    hasLeadingWhiteSpace = true; // VBScript: Set .ActiveConnection = oConn
                                }
                                else
                                {
                                    hasLeadingWhiteSpace = false;
                                }
                            }
                        }
                        else
                        {
                            hasLeadingWhiteSpace = false;// last_chr0IsWhitespace ?? false;
                        }

                        if (chr == "=" && tokens.Count > 0 && tokens[tokens.Count - 1] is ComparisonOperatorToken cmpToken && (cmpToken.ContentUpperX().UpperText == "<" || cmpToken.ContentUpperX().UpperText == ">"))
                        {
                            // replace ('<' with '<=') or ('>' with '>=')
                            tokens.RemoveRange(tokens.Count - 1, 1); // remove last
                            string cmpText = cmpToken.Content + "=";
                            tokens.Add(AtomToken.GetNewToken(new StringUpper(cmpText), hasLeadingWhiteSpace, cmpToken.LineIndex));
                        }
                        else
                        {
                            tokens.Add(AtomToken.GetNewToken(chr.ToUpperX(), hasLeadingWhiteSpace, lineIndex));
                        }
                        buffer = "";
                    }
                    else
                    {
                        buffer += chr;
                    }
                }
                if (chr == "\n")
                {
                    lineIndex++;
                }

                last_chr0IsWhitespace = chr0IsWhitespace;
            }// while
            if (buffer != "")
            {
                tokens.Add(AtomToken.GetNewToken(buffer.ToUpperX(), hasLeadingWhiteSpace: false, lineIndex));
            }
#pragma warning restore CA1820 // Test for empty strings using string length

            // Handle ignore-line-return / end-of-statement combinations
            tokens = handleLineReturnCancels(tokens);

            return tokens;
        }

        /// <summary>
        /// Look for any "_" character AtomTokens and ensure they are followed by a line return - if so, drop both (if not, raise exception - invalid VBScript)
        /// </summary>
        private static List<IToken> handleLineReturnCancels(List<IToken> tokens)
        {
            var tokensOut = new List<IToken>();
            for (int index = 0; index < tokens.Count; index++)
            {
                var token = tokens[index];
                if ((token is AtomToken) && (token.Content == "_"))
                {
                    // Ensure followed by line return, then ignore both tokens
                    if (index == (tokens.Count - 1))
                        throw new InvalidOperationException($"Encountered line-return cancellation that isn't followed by a line return - invalid. Line:{token.LineIndex}");
                    var tokenNext = tokens[index + 1];
                    if (!(tokenNext is EndOfStatementNewLineToken))
                        throw new InvalidOperationException($"Encountered line-return cancellation that isn't followed by a line return - invalid. Line:{token.LineIndex}. Next ({tokenNext.LineIndex} {tokenNext.GetType().Name}):{tokenNext.Content}");
                    index++;
                }
                else
                    tokensOut.Add(token);
            }
            return tokensOut;
        }
    }
}
