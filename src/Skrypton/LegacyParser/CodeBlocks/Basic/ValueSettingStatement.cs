using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbsX")]
    public enum ValueSetTypeOptions
    {
        [EnumMember] Let,
        [EnumMember] Set
    }

    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public sealed class ValueSettingStatement : IHaveNonNestedExpressions
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        /// <summary>
        /// This statement represents the setting of one value to the result of another expression, whether that be a fixed
        /// value, a variable's value or the return value of a method call
        /// </summary>
        public ValueSettingStatement(Expression valueToSet, Expression expression, ValueSetTypeOptions valueSetType)
        {
            if (valueToSet == null)
                throw new ArgumentNullException("valueToSet");
            if (expression == null)
                throw new ArgumentNullException("expression");
            if (!Enum.IsDefined(typeof(ValueSetTypeOptions), valueSetType))
                throw new ArgumentOutOfRangeException("valueSetType");

            ValueToSet = valueToSet;
            Expression = expression;
            ValueSetType = valueSetType;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null
        /// </summary>
        [DataMember] public Expression ValueToSet { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        [DataMember] public Expression Expression { get; private set; }

        [DataMember] public ValueSetTypeOptions ValueSetType { get; private set; }


        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
            get
            {
                yield return ValueToSet;
                yield return Expression;
            }
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            // The Statement class' GenerateBaseSource has logic about rendering strings of tokens and rules about whitespace around
            // (or not around) particular tokens, so the content from this class is wrapped up as a Statement so that the method may
            // be re-used without copying any of it here
            var assignmentOperator = AtomToken.GetNewToken("=".ToUpperX(), ValueToSet.Tokens.Last().LineIndex);
            var tokensList = ValueToSet.Tokens.Concat(new[] { assignmentOperator }).Concat(Expression.Tokens).ToList();
            if (ValueSetType == ValueSetTypeOptions.Set)
                tokensList.Insert(0, AtomToken.GetNewToken("Set".ToUpperX(), ValueToSet.Tokens.First().LineIndex));

            return (new Statement(tokensList, Statement.CallPrefixOptions.Absent)).GenerateBaseSource(indenter);
        }
    }
}
