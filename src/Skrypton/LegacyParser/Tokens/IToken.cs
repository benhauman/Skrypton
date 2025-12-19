using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.Tokens
{
    public interface IToken
    {
        string Content { get; }
        StringUpper ContentUpperX();

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        int LineIndex { get; }
    }
}
