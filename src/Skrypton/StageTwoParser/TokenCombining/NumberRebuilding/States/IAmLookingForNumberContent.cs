using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public interface IAmLookingForNumberContent
    {
        TokenProcessResult Process(IReadOnlyCollection<IToken> tokens, PartialNumberContent numberContent);
    }
}
