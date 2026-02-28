using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class CommentStatement : INonExecutableCodeBlock
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public CommentStatement(string content, int lineIndex)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (content.Contains("\n"))
                throw new ArgumentException("The content may not include any line returns");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndex));

            Content = content.TrimEnd();
            LineIndex = lineIndex;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null or contain any line returns. It may be blank and may have leading whitespace (though it won't have
        /// any trailing whitespace).
        /// </summary>
        [DataMember] public string Content { get; private set; }

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
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            if (indenter == null) throw new ArgumentNullException(nameof(indenter));
#pragma warning disable CA1820 // Test for empty strings using string length
            if (Content.Trim() == "")
                return "";
#pragma warning restore CA1820 // Test for empty strings using string length
            return indenter.Indent + "'" + Content;
        }
    }
}
