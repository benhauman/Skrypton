using System;
using System.Linq;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// VBScript has (little-known) support for escaping the names of variables, function and classes so that they can contain almost anything if wrapped in square
    /// brackets. This really is almost ANYTHING, other than closing square brackets and line returns since there is no escape character for this name-wrapping. So
    /// underscores may be used, which are not valid otherwise. So can names that begin with numbers, which are also not otherwise valid. Names may also contain
    /// quotes or whitespace, names may contain ONLY numbers, symbols and/or whitespace. The name itself can be blank since [] is valid. I've never seen this
    /// live in the wild, but it's valid nonetheless and this class is how names of that form are represented.
    /// </summary>
    [Serializable]
    public class EscapedNameToken : NameToken
    {
        public EscapedNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Allow, lineIndex)
        {
            // Note that blank or whitespace-only are acceptable for this content so we can only check for null here
            if (contentUpper.Length == 0)
                throw new ArgumentNullException("escapedContent");
            if (!contentUpper.Original.StartsWith("["))
                throw new ArgumentException("The content for an EscapedNameToken must start with an opening square bracket");
            if (!contentUpper.Original.EndsWith("]"))
                throw new ArgumentException("The content for an EscapedNameToken must end with a closing square bracket");
            if (contentUpper.Original.Count(c => c == ']') > 1)
                throw new ArgumentException("The content for an EscapedNameToken may only closing square bracket as the termination character, not within the content");
            if (contentUpper.Original.Any(c => c == '\n'))
                throw new ArgumentException("The content for an EscapedNameToken not contain any line returns");
        }

        public EscapedNameToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
    }
}
