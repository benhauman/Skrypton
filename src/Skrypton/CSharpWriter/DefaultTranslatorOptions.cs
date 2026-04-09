using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter
{
    internal sealed class DefaultTranslatorOptions : ITranslatorOptions
    {
        private readonly TranslationIssuesCollectorBase _issuesCollector;
        private readonly Dictionary<string, bool> _suppressions = new Dictionary<string, bool>();

        public DefaultTranslatorOptions(IReadOnlyCollection<string> suppressions, TranslationIssuesCollectorBase issuesCollector)
        {
            _issuesCollector = issuesCollector ?? throw new ArgumentNullException(nameof(issuesCollector));
            foreach (string suppression in suppressions.Select(x => x.ToUpperInvariant()).Distinct())
            {
                _suppressions.Add(suppression, true);
            }
        }

        public bool AcceptTranslationError(string errorKey) => _suppressions.TryGetValue(errorKey, out bool suppressed) && suppressed;
        public void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex)
        {
            if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
            _issuesCollector.UndeclaredNamedReferenceDetected(errorKey, referenceName, lineIndex);
        }
    }

    internal sealed class TranslationIssuesCollectorDefault : TranslationIssuesCollectorBase
    {
        internal static readonly TranslationIssuesCollectorDefault Instance = new TranslationIssuesCollectorDefault();
        private TranslationIssuesCollectorDefault()
        {

        }
        public override void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex)
        {
            // do nothing
        }
    }

    internal abstract class TranslationIssuesCollectorBase
    {
        public abstract void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex);
    }
}