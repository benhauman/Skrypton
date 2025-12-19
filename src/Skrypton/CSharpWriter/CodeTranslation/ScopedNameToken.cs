using System;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    public class ScopedNameToken : NameToken
    {
        public ScopedNameToken(StringUpper contentUpper, int lineIndex, ScopeLocationOptions scopeLocation) : base(contentUpper, lineIndex)
        {
            if (!Enum.IsDefined(typeof(ScopeLocationOptions), scopeLocation))
                throw new ArgumentOutOfRangeException("scopeLocation");

            ScopeLocation = scopeLocation;
        }
        public ScopedNameToken(string content, int lineIndex, ScopeLocationOptions scopeLocation) : this(content.ToUpperX(), lineIndex, scopeLocation) { } // test

        public ScopeLocationOptions ScopeLocation { get; private set; }
    }
}
