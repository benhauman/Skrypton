using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class ParsingExpression // Rename to 'ParsingExpression'
    {
        public ParsingExpression(IReadOnlyCollection<IExpressionSegment> segments)
        {
            if (segments == null)
                throw new ArgumentNullException(nameof(segments));

            Segments = segments.ToList().AsReadOnly();
            if (Segments.Count == 0)
                throw new ArgumentException("The segments set may not be empty");
            if (Segments.Any(t => t == null))
                throw new ArgumentException("Null reference encountered in segments set");
        }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        public IReadOnlyCollection<IExpressionSegment> Segments { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        public IEnumerable<IToken> AllTokens
        {
            get { return Segments.SelectMany(s => s.AllTokens); }
        }

        public string RenderedContent
        {
            get
            {
                return string.Join(
                    "",
                    Segments.Select(s => s.RenderedContent)
                );
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
