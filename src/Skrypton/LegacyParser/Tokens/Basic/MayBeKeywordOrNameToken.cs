using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// There are tokens that may be reference names or keywords, depending upon context - it is not possible to tell just from their content. Values
    /// (such as "step") are represented by these tokens.
    /// </summary>
    [Serializable]
    public class MayBeKeywordOrNameToken : NameToken
    {
        public MayBeKeywordOrNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            if (!AtomToken.isContextDependentKeywordUpper(contentUpper))
                throw new ArgumentException("Invalid content for a MayBeKeywordOrNameToken");
        }
    }
}
