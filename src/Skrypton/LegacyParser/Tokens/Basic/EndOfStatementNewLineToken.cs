using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class EndOfStatementNewLineToken : AbstractEndOfStatementToken
    {
        public EndOfStatementNewLineToken(int lineIndex) : base(lineIndex) { }

        public override string Content
        {
            get { return ""; }
        }
    }
}
