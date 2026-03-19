using Skrypton.LegacyParser.Tokens;
using System.Collections.Generic;

namespace Skrypton.StageTwoParser.TokenCombining.NumberRebuilding.States
{
    public interface IAmLookingForNumberContent
    {
        TokenProcessResult Process(IReadOnlyCollection<IToken> tokens, PartialNumberContent numberContent);
    }
}
