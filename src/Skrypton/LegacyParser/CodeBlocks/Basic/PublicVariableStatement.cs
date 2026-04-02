using System;
using System.Collections.Generic;
using System.Text;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    public class PublicVariableStatement : DimStatement
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
#pragma warning disable CA1002 // Do not expose generic lists
        public PublicVariableStatement(int lineIndex, List<DimVariable> variables) : base(lineIndex, variables) { }
#pragma warning restore CA1002 // Do not expose generic lists

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public override string GenerateBaseSource(IBaseSourceGenerationContext generationContext)
        {
            if (generationContext == null) throw new ArgumentNullException(nameof(generationContext));
            // Grab content from DimStatement..
            string baseContent = base.GenerateBaseSource(generationContext.NullIndenter());
            if ((baseContent == null)
            || (baseContent.Length < 4)
            || (!baseContent.Substring(0, 4).Equals("DIM ", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Unexpected content from base class");

            // .. and change to be ReDim (add in Preserve keyword, if required)
            StringBuilder output = new StringBuilder();
            output.Append(generationContext.Indenter.Indent);
            output.Append("Public ");
            output.Append(baseContent.Substring(4));
            return output.ToString();
        }
    }
}
