using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class InlineCommentToken : CommentToken
    {
        public InlineCommentToken(string content, int lineIndex) : base(content, lineIndex) { }
    }
}
