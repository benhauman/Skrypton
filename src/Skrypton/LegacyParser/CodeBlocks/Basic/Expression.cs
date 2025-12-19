using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public sealed class Expression : Statement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// An expression is code that evalutes to a value
        /// </summary>
        public Expression(IEnumerable<IToken> tokens) : base(tokens, CallPrefixOptions.Absent) { }
    }
}
