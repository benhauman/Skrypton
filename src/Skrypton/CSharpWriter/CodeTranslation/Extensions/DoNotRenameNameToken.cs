using System;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    /// <summary>
    /// This is a special derived class of NameToken, it will not be affected when passed through the GetMemberAccessTokenName extension method of a VBScriptNameRewriter
    /// (this may be useful when content is being injected into expressions to ensure that name rewriting isn't double-applied - it is used in the StatementTranslator,
    /// for example)
    /// </summary>
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class DoNotRenameNameToken : NameToken
    {
        public DoNotRenameNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Allow, lineIndex)
        {
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");
        }
        public DoNotRenameNameToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
    }
}
