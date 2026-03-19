using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    internal sealed class GotSomeIntegerNumberContent : IAmLookingForNumberContent
    {
        public static GotSomeIntegerNumberContent Instance { get { return new GotSomeIntegerNumberContent(); } }
        private GotSomeIntegerNumberContent() { }

        public TokenProcessResult Process(IReadOnlyCollection<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (numberContent == null)
                throw new ArgumentNullException(nameof(numberContent));

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            // The only continuation possibility for the number is if a decimal point is reached
            if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    [],
                    GotSomeDecimalNumberContent.Instance
                );
            }

            // If we're not at a decimal point then the end of the number content must have been reached
            // - Try to extract the number content so far and express that as a new token
            // - Return a "processedTokens" set of this and the current token (we don't need to worry about trying to process
            //   that here since it's not valid for two number tokens to exist adjacently with nothing in between)
            var numbericValueToken = numberContent.TryToExpressNumericValueTokenFromCurrentTokens();
            if (numbericValueToken == null)
                throw new InvalidOperationException("numberContent should describe a number, null was returned from TryToExpressNumberFromTokens - invalid content");
            return new TokenProcessResult(
                new PartialNumberContent(),
                new[] { numbericValueToken, token },
                CommonUtilities.GetDefaultProcessor(tokens)
            );
        }
    }
}
