using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Skrypton.LegacyParser.CodeBlocks.SourceRendering;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class SelectBlock : IHaveNestedContent, ICodeBlock
    {
        public SelectBlock(
            Expression expression,
            IEnumerable<CommentStatement> openingComments,
            IEnumerable<CaseBlockSegment> content)
        {
            if (openingComments == null)
                throw new ArgumentNullException(nameof(openingComments));
            if (content == null)
                throw new ArgumentNullException(nameof(content));

            OpeningComments = openingComments.ToArray();
            if (OpeningComments.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in openingComments set");

            Content = content.ToArray();
            if (Content.Any(c => c == null))
                throw new ArgumentException("Null reference encountered in content set");
            var firstUnsupportedContentSegment = Content.FirstOrDefault(c => !(c is CaseBlockExpressionSegment) && !(c is CaseBlockElseSegment));
            if (firstUnsupportedContentSegment != null)
                throw new ArgumentException("Unrecognised content element: " + firstUnsupportedContentSegment.GetType());
            if (((IEnumerable<CaseBlockSegment>)Content).Reverse().Skip(1).Any(c => c is CaseBlockElseSegment))
                throw new ArgumentException("Only the last content segment may be a CaseBlockElseSegment");

            Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null
        /// </summary>
        [DataMember] public Expression Expression { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set
        /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
        [DataMember] public CommentStatement[] OpeningComments { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set. All items will be CaseBlockExpressionSegment or
        /// CaseBlockElseSegment instances and only the last segment may be a CaseBlockElseSegment (note that it is valid in VBScript for the
        /// ONLY segment to be a CaseBlockElseSegment - in which case the select "Expression" will still be evaluated but the "Case Else"
        /// will always be entered)
        /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
        [DataMember] public CaseBlockSegment[] Content { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays

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
                return new ICodeBlock[] { Expression }
                    .Concat(Content.Select(c => c as CaseBlockExpressionSegment).Where(c => c != null).SelectMany(c => c!.Values))
                    .Concat(Content.SelectMany(c => c.Statements));
            }
        }

        // =======================================================================================
        // DESCRIPTION CLASSES
        // =======================================================================================
        [DataContract(Namespace = "http://vbs")]
#pragma warning disable CA1034 // Nested types should not be visible
        public abstract class CaseBlockSegment
#pragma warning restore CA1034 // Nested types should not be visible
        {
            protected CaseBlockSegment(IEnumerable<ICodeBlock> statements)
            {
                if (statements == null)
                    throw new ArgumentNullException(nameof(statements));

                Statements = statements.ToArray();
                if (Statements.Any(v => v == null))
                    throw new ArgumentException("Null reference encountered in statements set");
            }

            /// <summary>
            /// This will never be null nor contain any null references, but it may be an empty set
            /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
            [DataMember] public ICodeBlock[] Statements { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays
        }

        [DataContract(Namespace = "http://vbs")]
#pragma warning disable CA1034 // Nested types should not be visible
        public sealed class CaseBlockExpressionSegment : CaseBlockSegment
#pragma warning restore CA1034 // Nested types should not be visible
        {
            public CaseBlockExpressionSegment(IEnumerable<Expression> values, IEnumerable<ICodeBlock> statements) : base(statements)
            {
                if (values == null)
                    throw new ArgumentNullException(nameof(values));

                Values = values.ToArray();
                if (Values.Any(v => v == null))
                    throw new ArgumentException("Null reference encountered in openingComments set");
                if (Values.Length == 0)
                    throw new ArgumentException("values is an empty set  - invalid");
            }

            /// <summary>
            /// This will never be null, empty nor contain any null references
            /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
            [DataMember] public Expression[] Values { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays
        }

        [DataContract(Namespace = "http://vbs")]
#pragma warning disable CA1034 // Nested types should not be visible
        public sealed class CaseBlockElseSegment : CaseBlockSegment
#pragma warning restore CA1034 // Nested types should not be visible
        {
            public CaseBlockElseSegment(IEnumerable<ICodeBlock> statements) : base(statements) { }
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
            var output = new StringBuilder();

            output.Append(indenter.Indent + "SELECT CASE ");
            output.AppendLine(Expression.GenerateBaseSource(NullIndenter.Instance));

            if (OpeningComments != null)
            {
                foreach (CommentStatement statement in OpeningComments)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase()));
                output.AppendLine("");
            }

            for (int index = 0; index < Content.Length; index++)
            {
                // Render branch start
                CaseBlockSegment segment = Content.ElementAt(index);
                if (segment is CaseBlockExpressionSegment)
                {
                    output.Append(indenter.Increase().Indent);
                    output.Append("CASE ");
                    var valuesArray = ((CaseBlockExpressionSegment)segment).Values.ToArray();
                    for (int indexValue = 0; indexValue < valuesArray.Length; indexValue++)
                    {
                        Expression statement = valuesArray[indexValue];
                        output.Append(statement.GenerateBaseSource(NullIndenter.Instance));
                        if (indexValue < (valuesArray.Length - 1))
                            output.Append(", ");
                    }
                    output.AppendLine("");
                }
                else
                    output.AppendLine(indenter.Increase().Indent + "CASE ELSE");

                // Render branch content
                foreach (ICodeBlock statement in segment.Statements)
                    output.AppendLine(statement.GenerateBaseSource(indenter.Increase().Increase()));
            }

            output.Append(indenter.Indent + "END SELECT");
            return output.ToString();
        }
    }
}
