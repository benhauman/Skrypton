using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class InlineCommentStatement : CommentStatement
    {
        public InlineCommentStatement(string content, int lineIndex) : base(content, lineIndex) { }
    }
}
