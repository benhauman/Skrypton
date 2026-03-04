using System;
using System.Runtime.Serialization;

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

    [DataContract(Namespace = "http://vbs")]
    public sealed class WhiteSpaceToken : IToken
    {
        public WhiteSpaceToken(int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndex), "must be zero or greater");

            LineIndex = lineIndex;
            _contentUpper = new StringUpper(" ");
        }

        public string Content => " ";

        [NonSerialized] private readonly StringUpper _contentUpper;
        public StringUpper ContentUpperX()
        {
            return _contentUpper;
        }


        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int LineIndex { get; }
    }
}
