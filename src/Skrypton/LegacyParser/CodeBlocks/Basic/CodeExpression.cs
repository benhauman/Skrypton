using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class CodeExpression : Statement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// An codeExpression is code that evaluates to a value
        /// </summary>
        public CodeExpression(IReadOnlyCollection<IToken> tokens) : base(tokens, CallPrefixOptions.Absent) { }
    }
}
