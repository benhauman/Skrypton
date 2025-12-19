using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class ForBlock : ILoopOverNestedContent, ICodeBlock
    {
        public ForBlock(NameToken loopVar, Expression loopFrom, Expression loopTo, Expression loopStep, List<ICodeBlock> statements)
        {
            if (loopVar == null)
                throw new ArgumentNullException("loopVar");
            if (loopFrom == null)
                throw new ArgumentNullException("loopFrom");
            if (loopTo == null)
                throw new ArgumentNullException("loopTo");
            if (statements == null)
                throw new ArgumentNullException("statements");
            LoopVar = loopVar;
            LoopFrom = loopFrom;
            LoopTo = loopTo;
            LoopStep = loopStep;
            Statements = statements;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// It is not valid in VBScript for the loop variable to be anything other than a simple variable reference (it may be "i" but may not
        /// be "i(0)" or "i.Name", for example)
        /// </summary>
        [DataMember] public NameToken LoopVar { get; private set; }

        [DataMember] public Expression LoopFrom { get; private set; }

        [DataMember] public Expression LoopTo { get; private set; }

        /// <summary>
        /// Note: This may be null
        /// </summary>
        [DataMember] public Expression LoopStep { get; private set; }

        [DataMember] public IEnumerable<ICodeBlock> Statements { get; private set; }

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get
            {
                return new ICodeBlock[] { new Expression(new[] { LoopVar }), LoopFrom, LoopTo, LoopStep }
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
        public string GenerateBaseSource(SourceRendering.ISourceIndentHandler indenter)
        {
            StringBuilder output = new StringBuilder();

            // Open statement
            output.Append(indenter.Indent);
            output.Append("For ");
            output.Append(this.LoopVar.Content);
            output.Append(" = ");
            output.Append(this.LoopFrom.GenerateBaseSource(NullIndenter.Instance));
            output.Append(" To ");
            output.Append(this.LoopTo.GenerateBaseSource(NullIndenter.Instance));
            if (this.LoopStep != null)
            {
                output.Append(" Step ");
                output.Append(this.LoopStep.GenerateBaseSource(NullIndenter.Instance));
            }
            output.AppendLine("");

            // Render inner content
            foreach (ICodeBlock statement in this.Statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement
            output.Append(indenter.Indent + "Next");
            return output.ToString();
        }
    }
}
