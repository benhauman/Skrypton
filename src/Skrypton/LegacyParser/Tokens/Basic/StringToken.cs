using System;
using System.Diagnostics;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a single string section
    /// </summary>
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    [DebuggerDisplay("{Content}")]
    public sealed class StringToken : IToken
    {
        public StringToken(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            Content = content;
            LineIndex = lineIndex;
        }

        /// <summary>
        /// This will not include the quotes in the value
        /// </summary>
        [DataMember] public string Content { get; private set; }
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
        [DataMember] public int LineIndex { get; private set; }

        public override string ToString()
        {
            return base.ToString() + ":" + Content;
        }
    }
}
