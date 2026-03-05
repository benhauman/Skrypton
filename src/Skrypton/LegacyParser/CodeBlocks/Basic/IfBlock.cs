using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class IfBlock : IHaveNestedContent
    {
        // =======================================================================================
        // CLASS INITIALISATION
        // =======================================================================================
        public IfBlock(IEnumerable<IfBlockSegment> clauses)
        {
            if (clauses == null)
                throw new ArgumentNullException(nameof(clauses));

            IfBlockSegment[] clausesArray = clauses.ToArray();
            if (clausesArray.Length == 0)
                throw new ArgumentException("Empty clauses set specified - invalid");
            if (clausesArray.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in clauses set");

            int numberOfElseSegments = clausesArray.Count(c => c is IfBlockElseSegment);
            if (numberOfElseSegments > 1)
                throw new ArgumentException("There may never be more than one IfBlockElseSegment");
            if (numberOfElseSegments == 1)
            {
                if ((clausesArray.Length == 1) || !(clausesArray.Last() is IfBlockElseSegment))
                    throw new ArgumentException("If an IfBlockElseSegment is present, it must be the last clause (and is not allowed if there is only a single clause");
            }

            IfBlockSegment? firstInvalidSegmentIfAny = clausesArray.FirstOrDefault(c => !(c is IfBlockConditionSegment) && !(c is IfBlockElseSegment));
            if (firstInvalidSegmentIfAny != null)
                throw new ArgumentException("Unsupported segment type: " + firstInvalidSegmentIfAny.GetType());

            ConditionalClauses = clausesArray.OfType<IfBlockConditionSegment>();
            OptionalElseClause = (IfBlockElseSegment?)clausesArray.FirstOrDefault(c => c is IfBlockElseSegment);
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null, empty or contain any nulls
        /// </summary>
        [DataMember] public IEnumerable<IfBlockConditionSegment> ConditionalClauses { get; private set; }

        /// <summary>
        /// This will be null if there was no fallback clause
        /// </summary>
        [DataMember] public IfBlockElseSegment? OptionalElseClause { get; private set; }

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
                foreach (IfBlockConditionSegment? conditionalClause in ConditionalClauses)
                {
                    yield return conditionalClause.Condition;
                    foreach (ICodeBlock? statement in conditionalClause.Statements)
                        yield return statement;
                }
                if (OptionalElseClause == null)
                    yield break;
                foreach (ICodeBlock? statement in OptionalElseClause.Statements)
                    yield return statement;
            }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
#pragma warning disable CA1034 // Nested types should not be visible
#pragma warning disable CA1715 // Identifiers should have correct prefix
        public interface IfBlockSegment
#pragma warning restore CA1715 // Identifiers should have correct prefix
#pragma warning restore CA1034 // Nested types should not be visible
        {
            IEnumerable<ICodeBlock> Statements { get; }
        }

        [DataContract(Namespace = "http://vbs")]
#pragma warning disable CA1034 // Nested types should not be visible
        public sealed class IfBlockConditionSegment : IfBlockSegment
#pragma warning restore CA1034 // Nested types should not be visible
        {
            public IfBlockConditionSegment(CodeExpression conditionStatement, IEnumerable<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException(nameof(statements));

                Statements = statements.ToList().AsReadOnly();
                if (Statements.Any(s => s == null))
                    throw new ArgumentException("Null reference encountered in statements set");
                Condition = conditionStatement ?? throw new ArgumentNullException(nameof(conditionStatement));
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            [DataMember] public CodeExpression Condition { get; private set; }

            /// <summary>
            /// This will never be null or contain any nulls
            /// </summary>
            [DataMember] public IEnumerable<ICodeBlock> Statements { get; private set; }
        }

        [DataContract(Namespace = "http://vbs")]
#pragma warning disable CA1034 // Nested types should not be visible
        public sealed class IfBlockElseSegment : IfBlockSegment
#pragma warning restore CA1034 // Nested types should not be visible
        {
            public IfBlockElseSegment(IEnumerable<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException(nameof(statements));
                Statements = statements.ToList().AsReadOnly();
                if (Statements.Any(s => s == null))
                    throw new ArgumentException("Null reference encountered in statements set");
            }

            /// <summary>
            /// This will never be null or contain any nulls
            /// </summary>
            [DataMember] public IEnumerable<ICodeBlock> Statements { get; private set; }
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

            List<IfBlockSegment> allClauses = ConditionalClauses.Cast<IfBlockSegment>().ToList();
            if (OptionalElseClause != null)
            {
                allClauses.Add(OptionalElseClause);
            }

            for (int index = 0; index < allClauses.Count; index++)
            {
                // Render branch start: IF / ELSEIF / ELSE
                IfBlockSegment? segment = allClauses[index];
                if (segment is IfBlockConditionSegment ifSegment)
                {
                    output.Append(indenter.Indent);
                    if (index == 0)
                        output.Append("IF ");
                    else
                        output.Append("ELSEIF ");
                    output.Append(
                        ifSegment.Condition.GenerateBaseSource(NullIndenter.Instance)
                    );
                    output.AppendLine(" THEN");
                }
                else
                {
                    output.AppendLine(indenter.Indent + "ELSE");
                }

                // Render branch content
                foreach (ICodeBlock statement in segment.Statements)
                {
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
                }
            }
            output.Append(indenter.Indent + "END IF");
            return output.ToString();
        }
    }
}
