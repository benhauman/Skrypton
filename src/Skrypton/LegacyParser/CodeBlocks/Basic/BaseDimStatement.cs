using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public abstract class BaseDimStatement : IHaveNonNestedExpressions
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        protected BaseDimStatement(IEnumerable<DimVariable> variables)
        {
            if (variables == null)
                throw new ArgumentNullException(nameof(variables));

            Variables = variables.ToList().AsReadOnly();
            if (Variables.Any(v => v == null))
                throw new ArgumentException("Null reference encountered in variables set");
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null nor contain any nulls (though it may be an empty set)
        /// </summary>
        [DataMember] public IEnumerable<DimVariable> Variables { get; protected set; }

        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
#pragma warning disable CA1033 // Interface methods should be callable by child types
            get { return Variables.Where(v => v.Dimensions != null).SelectMany(v => v.Dimensions); }
#pragma warning restore CA1033 // Interface methods should be callable by child types
        }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public virtual string GenerateBaseSource(IBaseSourceGenerationContext generationContext)
        {
            if (generationContext == null) throw new ArgumentNullException(nameof(generationContext));
            var output = new StringBuilder();
            output.Append(generationContext.Indenter.Indent);
            output.Append("Dim ");
            var numberOfVariables = Variables.Count();
            foreach (var indexedVariable in Variables.Select((v, i) => new { Variable = v, Index = i }))
            {
                output.Append(indexedVariable.Variable.Name.Content);
                if (indexedVariable.Variable.Dimensions != null)
                {
                    output.Append('(');
                    var numberOfDimensions = indexedVariable.Variable.Dimensions.Length;
                    foreach (var indexedDimension in indexedVariable.Variable.Dimensions.Select((d, i) => new { Dimension = d, Index = i }))
                    {
                        output.Append(indexedDimension.Dimension.GenerateBaseSource(generationContext.NullIndenter()));
                        if (indexedDimension.Index < (numberOfDimensions - 1))
                            output.Append(", ");
                    }
                    output.Append(')');
                }
                if (indexedVariable.Index < (numberOfVariables - 1))
                    output.Append(", ");
            }
            return output.ToString();
        }
    }

    // =======================================================================================
    // DESCRIPTION CLASSES
    // =======================================================================================
    [DataContract(Namespace = "http://vbs")]
    public class DimVariable // cannot be abstract due to 'translateRawVariableData'
    {
        public DimVariable(NameToken name, IEnumerable<CodeExpression>? dimensions)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            if (dimensions == null)
            {
                Dimensions = null;
            }
            else
            {
                Dimensions = dimensions.ToArray();
                if (Dimensions.Any(d => d == null))
                    throw new ArgumentException("Null reference encountered in dimensions set");
            }
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        [DataMember] public NameToken Name { get; private set; }

        /// <summary>
        /// Variables list may be null (not explicitly defined as an array), have zero elements (an uninitialised array) or multiple dimensions (but
        /// if the list is non-null and non-empty, it will never contain any null references)
        /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
        [DataMember] public CodeExpression[]? Dimensions { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays

        public override string ToString()
        {
            return base.ToString() + ":" + Name;
        }
    }
}
