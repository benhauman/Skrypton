using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace Skrypton.LegacyParser.CodeBlocks.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class SelectBlock : IHaveNestedContent, ICodeBlock
    {
        public SelectBlock(
            CodeExpression codeExpression,
            IEnumerable<CommentStatement> openingComments,
            IReadOnlyCollection<CaseBlockSegment> content)
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
                throw new ArgumentException($"Only the last content segment may be a CaseBlockElseSegment. Line:{codeExpression?.Tokens.FirstOrDefault()?.LineIndex}"); // 'Case Else' must be the last one (after Case(se))

            Expression = codeExpression ?? throw new ArgumentNullException(nameof(codeExpression));
        }

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be null
        /// </summary>
        [DataMember] public CodeExpression Expression { get; private set; }

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set
        /// </summary>
#pragma warning disable CA1819 // Properties should not return arrays
        [DataMember] public CommentStatement[] OpeningComments { get; private set; }
#pragma warning restore CA1819 // Properties should not return arrays

        /// <summary>
        /// This will never be null nor contain any null references, but it may be an empty set. All items will be CaseBlockExpressionSegment or
        /// CaseBlockElseSegment instances and only the last segment may be a CaseBlockElseSegment (note that it is valid in VBScript for the
        /// ONLY segment to be a CaseBlockElseSegment - in which case the select "ParsingExpression" will still be evaluated but the "Case Else"
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
            public CaseBlockExpressionSegment(IEnumerable<CodeExpression> values, IEnumerable<ICodeBlock> statements) : base(statements)
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
            [DataMember] public CodeExpression[] Values { get; private set; }
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
        public string GenerateBaseSource(IBaseSourceGenerationContext generationContext)
        {
            if (generationContext == null) throw new ArgumentNullException(nameof(generationContext));
            var output = new StringBuilder();

            output.Append(generationContext.Indent + "SELECT CASE ");
            output.AppendLine(Expression.GenerateBaseSource(generationContext.NullIndenter()));

            if (OpeningComments.Length > 0)
            {
                foreach (CommentStatement statement in OpeningComments)
                    output.AppendLine(statement.GenerateBaseSource(generationContext.Increase()));
                output.AppendLine("");
            }

            for (int index = 0; index < Content.Length; index++)
            {
                // Render branch start
                CaseBlockSegment segment = Content.ElementAt(index);
                if (segment is CaseBlockExpressionSegment)
                {
                    output.Append(generationContext.Increase().Indent);
                    output.Append("CASE ");
                    var valuesArray = ((CaseBlockExpressionSegment)segment).Values.ToArray();
                    for (int indexValue = 0; indexValue < valuesArray.Length; indexValue++)
                    {
                        CodeExpression statement = valuesArray[indexValue];
                        output.Append(statement.GenerateBaseSource(generationContext.NullIndenter()));
                        if (indexValue < (valuesArray.Length - 1))
                            output.Append(", ");
                    }
                    output.AppendLine("");
                }
                else
                    output.AppendLine(generationContext.Increase().Indent + "CASE ELSE");

                // Render branch content
                foreach (ICodeBlock statement in segment.Statements)
                    output.AppendLine(statement.GenerateBaseSource(generationContext.Increase().Increase()));
            }

            output.Append(generationContext.Indent + "END SELECT");
            return output.ToString();
        }
    }
}
