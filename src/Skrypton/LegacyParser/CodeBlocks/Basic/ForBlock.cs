using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class ForBlock : ILoopOverNestedContent, ICodeBlock
    {
        public ForBlock(NameToken loopVar, CodeExpression loopFrom, CodeExpression loopTo, CodeExpression? loopStep, IList<ICodeBlock> statements)
        {
            LoopVar = loopVar ?? throw new ArgumentNullException(nameof(loopVar));
            LoopFrom = loopFrom ?? throw new ArgumentNullException(nameof(loopFrom));
            LoopTo = loopTo ?? throw new ArgumentNullException(nameof(loopTo));
            LoopStep = loopStep;
            Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// It is not valid in VBScript for the loop variable to be anything other than a simple variable reference (it may be "i" but may not
        /// be "i(0)" or "i.Name", for example)
        /// </summary>
        [DataMember] public NameToken LoopVar { get; private set; }

        [DataMember] public CodeExpression LoopFrom { get; private set; }

        [DataMember] public CodeExpression LoopTo { get; private set; }

        /// <summary>
        /// Note: This may be null
        /// </summary>
        [DataMember] public CodeExpression? LoopStep { get; private set; }

        [DataMember] public IEnumerable<ICodeBlock> Statements { get; private set; }

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
#pragma warning disable CA1033 // Interface methods should be callable by child types
            get
#pragma warning restore CA1033 // Interface methods should be callable by child types
            {
                return new ICodeBlock[] { new CodeExpression(new[] { LoopVar }), LoopFrom, LoopTo, LoopStep! }
                    .Where(b => b != null) // Ignore a null LoopStep (this is a valid configuration but we can't have nulls in the data returned here)
                    .Concat(Statements);
            }
        }

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

            // Open statement
            output.Append(generationContext.Indent);
            output.Append("For ");
            output.Append(this.LoopVar.Content);
            output.Append(" = ");
            output.Append(this.LoopFrom.GenerateBaseSource(generationContext.NullIndenter()));
            output.Append(" To ");
            output.Append(this.LoopTo.GenerateBaseSource(generationContext.NullIndenter()));
            if (this.LoopStep != null)
            {
                output.Append(" Step ");
                output.Append(this.LoopStep.GenerateBaseSource(generationContext.NullIndenter()));
            }
            output.AppendLine("");

            // Render inner content
            foreach (ICodeBlock statement in this.Statements)
                output.AppendLine(statement.GenerateBaseSource(generationContext.Increase()));

            // Close statement
            output.Append(generationContext.Indent + "Next");
            return output.ToString();
        }
    }
}
