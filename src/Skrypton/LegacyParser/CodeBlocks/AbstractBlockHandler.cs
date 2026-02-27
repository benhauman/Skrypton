using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks
{
    public abstract class AbstractBlockHandler // public due to tests
    {
        // =======================================================================================
        // ABSTRACT METHODS
        // =======================================================================================
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public abstract ICodeBlock Process(List<IToken> tokens);

        // =======================================================================================
        // HELPER METHODS FOR DERIVED CLASSES
        // =======================================================================================
        /// <summary>
        /// Grab specific token from list. Optionally specify that it must be an AtomToken in
        /// order to be valid. Will raise an exception if there are no more tokens available,
        /// or if a AtomToken was required but the next token was of a different type.
        /// </summary>
        protected static IToken getToken(IReadOnlyCollection<IToken> tokens, int offset, IReadOnlyCollection<Type> allowedTokenTypes)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (offset < 0)
                throw new ArgumentException("Negative offset specified - invalid");
            if (offset >= tokens.Count())
                throw new ArgumentException("Insufficient tokens - invalid");
            if ((allowedTokenTypes != null) && allowedTokenTypes.Count == 0)
                throw new ArgumentException("No allowed tokens types (pass as null to set no restriction");
            var token = tokens.ElementAt(offset);
            if (allowedTokenTypes != null)
            {
                bool validTokenType = false;
                foreach (var allowedType in allowedTokenTypes)
                {
                    if (isObjectOfTypeOrDerivedFrom(token, allowedType))
                    {
                        validTokenType = true;
                        break;
                    }
                }
                if (!validTokenType)
                    throw new InvalidOperationException("Token is not of an allowed type [" + token.GetType().ToString() + " on line " + (token.LineIndex + 1) + "]");
            }
            return token;
        }

        private static bool isObjectOfTypeOrDerivedFrom(object obj, Type type)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));
            if (type == null)
                throw new ArgumentNullException(nameof(type));
            var objType = obj.GetType();
            while (true)
            {
                if (objType == type)
                    return true;
                if (objType.BaseType == null)
                    return false;
                objType = objType.BaseType;
            }
        }

        protected static IToken getToken_AtomOnly(IReadOnlyCollection<IToken> tokens, int offset)
        {
            return getToken(tokens, offset, new List<Type>()
            {
                typeof(AtomToken)
            });
        }

        protected static IToken getToken_AtomOrDateStringLiteralOnly(IReadOnlyCollection<IToken> tokens, int offset)
        {
            return getToken(tokens, offset, new List<Type>()
            {
                typeof(AtomToken),
                typeof(DateLiteralToken),
                typeof(StringToken)
            });
        }

        protected static bool isEndOfStatement(IReadOnlyCollection<IToken> tokens, int offset)
        {
            var token = getToken(tokens, offset, null);
            return (token is AbstractEndOfStatementToken);
        }

        /// <summary>
        /// Try to match AtomToken pattern - if there are insufficient tokens to match, or if a
        /// non-AtomToken is encountered, return false. Only rase exceptions if null tokens are
        /// found in the stream, the stream is null, the "values" array is null or empty of the
        /// optional offset value is less than zero. (If the offset value is too far along for
        /// the content to be matched, false will be returned).
        /// </summary>
        protected static bool checkAtomTokenPattern(IReadOnlyCollection<IToken> tokens, string[] values, bool matchCase)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (values == null)
                throw new ArgumentNullException(nameof(values));
            if (values.Length == 0)
                throw new ArgumentException("Zero values to match");

            var tokensToConsider = tokens.Take(values.Length).ToArray();
            if (tokensToConsider.Count() < values.Length)
                return false;

            var index = 0;
            foreach (var token in tokensToConsider)
            {
                if (token == null)
                    throw new ArgumentException("Null token specified");

                var value = values[index];
                if (value == null)
                    throw new ArgumentException("Null reference encountered in values set");

                // Only consider AtomTokens (if get anything else, we can't handle it)
                if (!(token is AtomToken))
                    return false;

                if (!value.Equals(token.Content, matchCase ? StringComparison.CurrentCulture : StringComparison.OrdinalIgnoreCase))
                    return false;

                index++;
            }
            return true;
        }
        protected static bool checkAtomTokenPattern(IReadOnlyCollection<IToken> tokens, string matchPattern, bool matchCase)
        {
            return checkAtomTokenPattern(tokens, [matchPattern], matchCase);
        }

        protected static bool checkAtomTokenPattern(IReadOnlyCollection<IToken> tokens, int offset, string[] matchPatterns, bool matchCase)
        {
            // Validate input - throw exception if conditions not met
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (matchPatterns == null)
                throw new ArgumentNullException(nameof(matchPatterns));
            if (matchPatterns.Length == 0)
                throw new ArgumentException("Zero matchPatterns to match", nameof(matchPatterns));
            if (offset < 0)
                throw new ArgumentException("Invalid offset value < 0 [" + offset.ToString(CultureInfo.InvariantCulture) + "]", nameof(offset));

            // If there are insufficient tokens, return false rather than throwing an exception (this method is supposed to be flexible)
            return checkAtomTokenPattern(tokens.Skip(offset).ToArray(), matchPatterns, matchCase);
        }

        /// <summary>
        /// Extract a comma-separated sequence of values from a token stream, starting at the
        /// specified location. Continue until find a token matching the endMarker (both the
        /// token type and content must match). The endMarker will only be checked for after
        /// validated content - eg. if FunctionHandler needs to traverse parameter tokens with
        /// a ")" AtomToken endMarker, any "(", ")" sequences that complement each other will
        /// not count towards the endMarker. Only AtomTokens and StringTokens are permissible
        /// in the token stream that are to be handled here, with the exception of the optional
        /// use of an EndOfStatementToken for the endMarker.
        /// </summary>
        protected static List<List<IToken>> getEntryList(IReadOnlyCollection<IToken> tokens, int offset, IToken endMarker)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (offset < 0)
                throw new InvalidOperationException("Negative offset specified - invalid");
            if (offset >= tokens.Count())
                throw new InvalidOperationException("Insufficient tokens - invalid");
            if (endMarker == null)
                throw new ArgumentNullException(nameof(endMarker));
            if ((!(endMarker is AtomToken)) && (!(endMarker is AbstractEndOfStatementToken)))
                throw new ArgumentException("Invalid endMarker - must be Atom or EndOfStatement Token");

            var allowedTokenTypes = new List<Type>() { typeof(AtomToken), typeof(DateLiteralToken), typeof(StringToken) };
            var entryList = new List<List<IToken>>();
            var buffer = new List<IToken>();
            var bracketCount = 0;
            while (true)
            {
                // Only check for endMarker if not in bracket sequence
                if (bracketCount == 0)
                {
                    // Check for endMarker
                    bool reachedEndMarker = false;
                    if ((offset >= tokens.Count()) || (endMarker is AbstractEndOfStatementToken) && isEndOfStatement(tokens, offset))
                        reachedEndMarker = true;
                    else
                    {
                        var possibleEndMarker = getToken(tokens, offset, allowedTokenTypes);
                        reachedEndMarker =
                            ((possibleEndMarker is AtomToken)
                            && (possibleEndMarker.Content.Equals(endMarker.Content, StringComparison.OrdinalIgnoreCase)));
                    }
                    if (reachedEndMarker)
                        break;
                }

                // Only check for separator if not in bracket sequence
                var gotSeparator = false;
                if (bracketCount == 0)
                {
                    IToken token = getToken(tokens, offset, allowedTokenTypes);
                    if (token is ArgumentSeparatorToken)
                    {
                        // Got it.. add current entry to list (don't worry if it's blank,
                        // let the caller decide whether that's valid or not)
                        gotSeparator = true;
                        entryList.Add(buffer);
                        buffer = new List<IToken>();
                    }
                }
                if (!gotSeparator)
                {
                    // Not got separator, add to buffer (check for brackets)
                    IToken token = getToken(tokens, offset, allowedTokenTypes);
                    buffer.Add(token);
                    if (token is OpenBrace)
                        bracketCount++;
                    else if (token is CloseBrace)
                    {
                        bracketCount--;
                        if (bracketCount < 0)
                            throw new ArgumentException("Mismatched brackets on ERASE statement on line " + (token.LineIndex + 1));
                    }
                }
                offset++;
            }
            if (buffer.Count != 0)
                entryList.Add(buffer);
            return entryList;
        }

        /// <summary>
        /// Return a new list that is a subset of the input token list
        /// </summary>
        protected static List<IToken> getTokenListSection(IReadOnlyCollection<IToken> tokens, int start, int count)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));

            var numberOfTokens = tokens.Count();
            if ((start < 0) || (start >= numberOfTokens))
                throw new ArgumentException(FormattableString.Invariant($"Invalid start value [{start}]"));
            if ((count < 0) || (start + count > numberOfTokens))
                throw new ArgumentException(FormattableString.Invariant($"Invalid count value [{start},{count}]"));

            return tokens.Skip(start).Take(count).ToList();
        }

        /// <summary>
        /// Return a new list that is a subset of the input token list - taken from the start position to the end of the token list
        /// </summary>
        protected static List<IToken> getTokenListSection(IReadOnlyCollection<IToken> tokens, int start)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));

            return getTokenListSection(tokens, start, tokens.Count() - start);
        }
    }
}