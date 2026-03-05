using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class OnErrorResumeNext : ICodeBlock
    {
        public OnErrorResumeNext(int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndex));

            LineIndex = lineIndex;
        }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        [DataMember] public int LineIndex { get; private set; }

        // =======================================================================================
        // VBScript BASE SOURCE RE-GENERATION
        // =======================================================================================
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        public string GenerateBaseSource(IBaseSourceGenerationContext generationContext)
        {
            if (generationContext == null) throw new ArgumentNullException(nameof(generationContext));
            return generationContext.Indent + "ON ERROR RESUME NEXT";
        }
    }
}
