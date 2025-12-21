using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public sealed class KeyWordToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public KeyWordToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");
            if (!AtomToken.isMustHandleKeyWordUpper(contentUpper) && !AtomToken.isContextDependentKeywordUpper(contentUpper) && !AtomToken.isMiscKeyWordUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not a VBScript keyword");
        }
        public KeyWordToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
    }
}
