using System;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    /// <summary>
    /// Sometimes an expression needs to be rewritten after some of it has been processed such that, not only must it not be pushed through the name rewriter again, the scope
    /// of the target reference is already known and so the logic that tries to determine whether it is a global, environment, local or undeclared reference may be avoided.
    /// </summary>
    [Serializable]
    public class ProcessedNameToken : DoNotRenameNameToken
    {
        public ProcessedNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, lineIndex)
        {
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");
        }
    }
}
