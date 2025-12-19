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
    public class ForEachBlock : ILoopOverNestedContent, ICodeBlock
    {
        /// <summary>
        /// It is valid to have a null conditionStatement in VBScript - in case the
        /// doUntil value is not of any consequence
        /// </summary>
        public ForEachBlock(NameToken loopVar, Expression loopSrc, List<ICodeBlock> statements)
        {
            if (loopVar == null)
                throw new ArgumentNullException("loopVar");
            if (loopSrc == null)
                throw new ArgumentNullException("loopSrc");
            if (statements == null)
                throw new ArgumentNullException("statements");
            this.LoopVar = loopVar;
            this.LoopSrc = loopSrc;
            this.Statements = statements;
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        [DataMember] public NameToken LoopVar { get; private set; }

        [DataMember] public Expression LoopSrc { get; private set; }

        [DataMember] public List<ICodeBlock> Statements { get; private set; }

        /// <summary>
        /// This is a flattened list of executable statements - for a function this will be the statements it contains but for an if block it
        /// would include the statements inside the conditions but also the conditions themselves. It will never be null nor contain any nulls.
        /// Note that this does not recursively drill down through nested code blocks so there will be cases where there are more executable
        /// blocks within child code blocks.
        /// </summary>
        IEnumerable<ICodeBlock> IHaveNestedContent.AllExecutableBlocks
        {
            get { return new ICodeBlock[] { new Expression(new[] { LoopVar }), LoopSrc }.Concat(Statements); }
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
            output.Append("For Each ");
            output.Append(this.LoopVar.Content);
            output.Append(" In ");
            output.AppendLine(this.LoopSrc.GenerateBaseSource(NullIndenter.Instance));

            // Render inner content
            foreach (ICodeBlock statement in this.Statements)
                output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));

            // Close statement
            output.Append(indenter.Indent + "Next");
            return output.ToString();
        }
    }
}
