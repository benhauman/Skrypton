/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBScriptTranslator.LegacyParser
{
    public interface ICodeBlock : IFragment
    {
        /// <summary>
        /// Re-generate equivalent VBScript source code for this block - there
        /// should not be a line return at the end of the content
        /// </summary>
        //string GenerateBaseSource(SourceRendering.IBaseSourceGenerator generator, SourceRendering.ISourceIndentHandler indenter);
        void AddInlineComment(ICodeBlock commentBlock); // CommentStatement

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        int LineIndex { get; }

    }
}
*/