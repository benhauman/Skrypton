using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks
{
    internal sealed class CodeBlockHandler
    {
        private readonly IReadOnlyCollection<string[]>? blockEnds;
        internal static readonly CodeBlockHandler RootBlock = new CodeBlockHandler(true, null);

        public CodeBlockHandler(string[] optionalBlockEndsOrNull) : this(false, [optionalBlockEndsOrNull]) { }

        public CodeBlockHandler(IEnumerable<string[]> optionalBlockEndsOrNull) : this(false, optionalBlockEndsOrNull) { }
        private CodeBlockHandler(bool root, IEnumerable<string[]>? optionalBlockEndsOrNull)
        {
            // Null blockEnds value is valid - for the root code block - but if a non-null
            // list IS specified then every sequence must be non-null, have at least one
            // element and each token in the end sequences must be non-null
            if (root && optionalBlockEndsOrNull == null)
            {
                this.blockEnds = null;
                return;
            }
            if (optionalBlockEndsOrNull == null) throw new ArgumentNullException(nameof(optionalBlockEndsOrNull));

            var blockEnds = new List<string[]>();
            foreach (string[]? endSequence in optionalBlockEndsOrNull)
            {
                if (endSequence == null || endSequence!.Length == 0)
                    throw new ArgumentException("Invalid BlockEnd sequence specified: null or blank");
                if (endSequence.Any(s => s == null))
                    throw new ArgumentException("Invalid BlockEnd sequence specified: null value within sequence");
                blockEnds.Add(endSequence.ToArray()); // clone
            }
            this.blockEnds = blockEnds;
        }

        /// <summary>
        /// Try to process the token stream using the AbstractBlockHandler classes - this will raise an exception if none of the handlers can process
        /// the token stream at any point. If it processes the entire stream, the output parameter endSequenceMet will be set as null. If the processing
        /// ends when an end sequence (defined at class initialisation) is met, that sequence will be passed out. This will never return null. Note that
        /// the token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated.
        /// </summary>
        public List<ICodeBlock> Process(List<IToken> tokens, out string[]? endSequenceMet)
        {
            if (tokens == null) throw new ArgumentNullException(nameof(tokens));
            var handlers = new AbstractBlockHandler[]
            {
                new OptionExplicitHandler(),

                new NewInstanceHandler(),

                new ClassHandler(),
                new FunctionHandler(),

                // This needs to come after FunctionHandler due to the Private/Public keyword overlap
                new DimHandler(),
                new EraseHandler(),
                new ConstHandler(),

                new DoHandler(),
                new WhileHandler(),

                new ForHandler(),

                new IfHandler(),
                new SelectHandler(),

                new ExitHandler(),

                new OnErrorHandler(),

                new RandomizeHandler(),

                new NoSpecialCaseHandler(),

                new WithHandler(),

                // This should always be used as a last resort
                new StatementHandler()
            };

            var codeBlocks = new List<ICodeBlock>();
            while (tokens.Count > 0)
            {
                // Check for sequence-end for current code block
                if (atBlockEnd(tokens, out endSequenceMet))
                    return codeBlocks;

                // If not sequence-end, try to process content
                ICodeBlock? codeBlock = null;
                foreach (var handler in handlers)
                {
                    codeBlock = handler.Process(tokens);
                    if (codeBlock != null)
                    {
                        codeBlocks.Add(codeBlock);
                        break;
                    }
                }
                if (codeBlock == null)
                    throw new InvalidOperationException("No handler able to handle token stream at current position");
            }
            endSequenceMet = null;
            return codeBlocks;
        }

        /// <summary>
        /// Have we encountered a token sequence that marks the end for this code block? Only KeyWordTokens are considered since this should be looking
        /// for structural termination sequences.
        /// </summary>
        private bool atBlockEnd(List<IToken> tokens, out string[]? endSequenceMet)
        {
            if (this.blockEnds != null)
            {
                foreach (string[] endSequence in this.blockEnds)
                {
                    if (tokens.Count >= endSequence.Length)
                    {
                        int matchCount = 0;
                        for (int index = 0; index < endSequence.Length; index++)
                        {
                            var token = tokens[index];
                            bool isTokenThatNeedsCheckingAgainstEndSequence;
                            if (token is KeyWordToken)
                                isTokenThatNeedsCheckingAgainstEndSequence = true;
                            else if ((index > 0) && (token is MayBeKeywordOrNameToken))
                            {
                                // The token "Property" may be a keyword (as most commonly expected) or a reference name (eg. "property = 1" is valid).
                                // So "Property" is categorised as a MayBeKeywordOrNameToken, rather than KeyWordToken. This means that when looking
                                // for "End" / "Property", we can't only consider KeyWordTokens. We can't just allow matching MayBeKeywordOrNameToken
                                // unless a false positive is found for a token that is actually a reference name, so the first token must be a
                                // KeywordToken and subsequent tokens may be either a KeywordToken or MayBeKeywordOrNameToken. This seems like
                                // a safe compromise (an alternative may be to allow MayBeKeywordOrNameToken for any token, for end sequence
                                // matches that are more than one token long - but I'll leave this approach unless it proves problematic).
                                isTokenThatNeedsCheckingAgainstEndSequence = true;
                            }
                            else
                                isTokenThatNeedsCheckingAgainstEndSequence = false;
                            if (!isTokenThatNeedsCheckingAgainstEndSequence)
                                continue;
                            if (token.Content.Equals(endSequence[index], StringComparison.OrdinalIgnoreCase))
                                matchCount++;
                        }
                        if (matchCount == endSequence.Length)
                        {
                            endSequenceMet = endSequence;
                            return true;
                        }
                    }
                }
            }
            endSequenceMet = null;
            return false;
        }
    }
}
