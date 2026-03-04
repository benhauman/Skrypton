using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class Expression : Statement // TODO: Rename To CodeExpression to reduce collisions with 'Skrypton.StageTwoParser.ExpressionParsing.Expression' and  'System.Linq.Expressions.Expression'
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// An expression is code that evalutes to a value
        /// </summary>
        public Expression(IReadOnlyCollection<IToken> tokens) : base(tokens, CallPrefixOptions.Absent) { }
    }
}
