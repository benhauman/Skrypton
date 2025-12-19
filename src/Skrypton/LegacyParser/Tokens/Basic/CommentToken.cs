using System;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public class CommentToken : IToken
    {
        public CommentToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            LineIndex = lineIndex;
            Content = content;
        }

        public string Content { get; private set; }
        [NonSerialized] StringUpper contentUpper;
        public StringUpper ContentUpperX()

        {
            if (contentUpper == null)
                contentUpper = Content.ToUpperX();
            return contentUpper;
        }


        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; private set; }
    }
}
