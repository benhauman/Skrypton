using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    public sealed class ClassBlock : ICodeBlock, IDefineScope
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        private NameToken className;
        private List<ICodeBlock> statements;
#pragma warning disable CA1002 // Do not expose generic lists
        public ClassBlock(NameToken className, List<ICodeBlock> statements)
#pragma warning restore CA1002 // Do not expose generic lists
        {
            if (statements == null)
                throw new ArgumentNullException(nameof(statements));

            foreach (ICodeBlock block in statements)
            {
                if (block == null)
                    throw new ArgumentException("Null block in statements");
            }

            this.className = className ?? throw new ArgumentNullException(nameof(className));
            this.statements = statements;
        }

        public override string ToString()
        {
            return base.ToString() + ":" + this.className.Content;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        public NameToken Name
        {
            get { return this.className; }
        }

        public IList<ICodeBlock> Statements
        {
            get { return this.statements; }
        }

        /// <summary>
        /// This must never be null but it may be empty (this may be the names of a a function's arguments, for example)
        /// </summary>
#pragma warning disable CA1033 // Interface methods should be callable by child types
        IEnumerable<NameToken> IDefineScope.ExplicitScopeAdditions => [];
#pragma warning restore CA1033 // Interface methods should be callable by child types

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
#pragma warning disable CA1033 // Interface methods should be callable by child types
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks => this.statements.AsReadOnly();

        ScopeLocationOptions IDefineScope.Scope { get { return ScopeLocationOptions.WithinClass; } }
#pragma warning restore CA1033 // Interface methods should be callable by child types

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
            StringBuilder output = new StringBuilder();
            output.AppendLine(generationContext.Indent + "Class " + this.className.Content);
            foreach (ICodeBlock block in this.statements)
                output.AppendLine(block.GenerateBaseSource(generationContext.Increase()));
            output.Append(generationContext.Indent + "End Class");
            return output.ToString();
        }
    }
}
