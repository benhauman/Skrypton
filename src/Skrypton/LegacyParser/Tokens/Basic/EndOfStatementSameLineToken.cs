using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public sealed class EndOfStatementSameLineToken : AbstractEndOfStatementToken
    {
        public EndOfStatementSameLineToken(int lineIndex) : base(lineIndex) { }

        public override string Content
        {
            get { return ""; }
        }
    }
}
