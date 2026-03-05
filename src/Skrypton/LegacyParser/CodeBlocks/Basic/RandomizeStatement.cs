using System;
using System.Collections.Generic;
using System.Text;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    internal sealed class RandomizeStatement : IHaveNonNestedExpressions
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public RandomizeStatement(int lineIndex, CodeExpression? seedIfAny)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndex));

            LineIndex = lineIndex;
            SeedIfAny = seedIfAny;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public int LineIndex { get; }

        /// <summary>
        /// Note: This may be null
        /// </summary>
		public CodeExpression? SeedIfAny { get; }

        /// <summary>
        /// This must never return null nor a set containing any nulls, it represents all executable statements within this structure that wraps statement(s)
        /// in a non-hierarhical manner (unlike the IfBlock, for example, which implements IHaveNestedContent rather than IHaveNonNestedExpressions)
        /// </summary>
        IEnumerable<Statement> IHaveNonNestedExpressions.NonNestedExpressions
        {
#pragma warning disable CA1033 // Interface methods should be callable by child types
            get
#pragma warning restore CA1033 // Interface methods should be callable by child types
            {
                if (SeedIfAny != null)
                    yield return SeedIfAny;
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
            if (indenter == null) throw new ArgumentNullException(nameof(indenter));
            StringBuilder output = new StringBuilder();
            output.Append(indenter.Indent);
            output.Append("Randomize");
            if (SeedIfAny != null)
            {
                output.Append(' ');
                output.Append(SeedIfAny.GenerateBaseSource(NullIndenter.Instance));
            }
            return output.ToString();
        }
    }
}
