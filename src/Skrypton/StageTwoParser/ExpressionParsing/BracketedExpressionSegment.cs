using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class BracketedExpressionSegment : IExpressionSegment
    {
        private readonly ReadOnlyCollection<IToken> _allTokens;
        public BracketedExpressionSegment(IReadOnlyCollection<IExpressionSegment> segments)
        {
            if (segments == null)
                throw new ArgumentNullException(nameof(segments));

            Segments = segments.ToList().AsReadOnly();
            if (Segments.Any(e => e == null))
                throw new ArgumentException("Null reference encountered in segments set");
            if (Segments.Count == 0)
                throw new ArgumentException("Empty segments set specified - invalid");

            // 2015-03-23 DWR: For deeply-nested bracketed segments, it can be very expensive to enumerate over their AllTokens sets repeatedly so it's worth preparing the data once and
            // avoiding doing it over and over again. This is often seen with an codeExpression with many string concatenations - currently they are broken down into pairs of operations,
            // which results in many bracketed operations (I want to change this for concatenations going forward, since it's so common to have sets of concatenations and it would
            // be better if the CONCAT took a variable number of arguments rather than just two, but this hasn't been done yet).
            _allTokens =
                new IToken[] { new OpenBrace(Segments.First().AllTokens.First().LineIndex) }
                .Concat(Segments.SelectMany(s => s.AllTokens))
                .Concat(new[] { new CloseBrace(Segments.Last().AllTokens.Last().LineIndex) })
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
		/// This will never be null, empty or contain any null references
		/// </summary>
		public IReadOnlyCollection<IExpressionSegment> Segments { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
#pragma warning disable CA1033 // Interface methods should be callable by child types
        IEnumerable<IToken> IExpressionSegment.AllTokens { get { return _allTokens; } }
#pragma warning restore CA1033 // Interface methods should be callable by child types

        public string RenderedContent
        {
            get
            {
                return "(" + string.Join("", Segments.Select(e => e.RenderedContent)) + ")";
            }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
