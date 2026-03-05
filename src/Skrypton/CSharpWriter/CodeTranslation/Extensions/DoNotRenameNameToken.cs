using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    /// <summary>
    /// This is a special derived class of NameToken, it will not be affected when passed through the GetMemberAccessTokenName extension method of a VBScriptNameRewriter
    /// (this may be useful when content is being injected into expressions to ensure that name rewriting isn't double-applied - it is used in the StatementTranslator,
    /// for example)
    /// </summary>
    [DataContract(Namespace = "http://vbs")]
    public class DoNotRenameNameToken : NameToken // not sealed due to 'ProcessedNameToken'
    {
        public DoNotRenameNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Allow, lineIndex)
        {
            if (contentUpper == null) throw new ArgumentNullException(nameof(contentUpper));
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");

            //if (!KnownDoNotRenameNames.TryGetValue(contentUpper.UpperText, out var isKnown))
            //{
            //    //throw new ArgumentException("Unknown name:" + contentUpper.UpperText, nameof(contentUpper));
            //}
        }
        public DoNotRenameNameToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test

        //private static readonly Dictionary<string, bool> KnownDoNotRenameNames = new Dictionary<string, bool>()
        //{
        //    {"ROOT", false},
        //    {"RUNNER", false},
        //    {"WITH", false}
        //};
    }
}
