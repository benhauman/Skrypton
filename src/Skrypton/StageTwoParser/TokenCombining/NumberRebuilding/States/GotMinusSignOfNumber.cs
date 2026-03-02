using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    internal sealed class GotMinusSignOfNumber : IAmLookingForNumberContent
    {
        public static GotMinusSignOfNumber Instance { get { return new GotMinusSignOfNumber(); } }
        private GotMinusSignOfNumber() { }

        public TokenProcessResult Process(IReadOnlyCollection<IToken> tokens, PartialNumberContent numberContent)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (numberContent == null)
                throw new ArgumentNullException(nameof(numberContent));

            var token = tokens.First();
            if (token == null)
                throw new ArgumentException("Null reference encountered in tokens set");

            // At this point, the current token needs to be either a number or a decimal point. Otherwise it's not going to be
            // part of a valid numeric value.
            if (token is NumericValueToken)
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    [],
                    GotSomeIntegerNumberContent.Instance
                );
            }
            else if (token.Is<MemberAccessorOrDecimalPointToken>())
            {
                return new TokenProcessResult(
                    numberContent.AddToken(token),
                    [],
                    GotSomeDecimalNumberContent.Instance
                );
            }
            return CommonUtilities.Reset(tokens, numberContent);
        }
    }
}
