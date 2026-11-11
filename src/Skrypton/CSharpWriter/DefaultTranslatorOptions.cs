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
        private readonly bool _undeclaredVariablesAreErrors;
        private bool _optionExplicit;

        /// <summary>
        /// undeclaredVariablesAreErrors: true (Skrypton's default) - every undeclared variable is an error (SKY101..SKY109) unless suppressed;
        /// false - VBScript's rule: undeclared variables are implicitly declared unless the script has "Option Explicit" (issue #86)
        /// </summary>
        public DefaultTranslatorOptions(IReadOnlyCollection<string> suppressions, TranslationIssuesCollectorBase issuesCollector, bool undeclaredVariablesAreErrors = true)
        {
            _issuesCollector = issuesCollector ?? throw new ArgumentNullException(nameof(issuesCollector));
            foreach (string suppression in suppressions.Select(x => x.ToUpperInvariant()).Distinct())
            {
                _suppressions.Add(suppression, true);
            }
            _undeclaredVariablesAreErrors = undeclaredVariablesAreErrors;
        }

        public bool AcceptTranslationError(string errorKey)
        {
            if (_suppressions.TryGetValue(errorKey, out bool suppressed) && suppressed)
                return true;
            // SKY101..SKY109: an undeclared variable - allowed with VBScript's rule unless the script has Option Explicit
            bool isUndeclaredVariable = errorKey.StartsWith("SKY10", StringComparison.Ordinal) && (errorKey.Length == 6) && (errorKey[5] >= '1') && (errorKey[5] <= '9');
            return isUndeclaredVariable && !_undeclaredVariablesAreErrors && !_optionExplicit;
        }

        public void OptionExplicitDetected() => _optionExplicit = true;

        public bool UsesVBScriptRuleForUndeclaredVariables => !_undeclaredVariablesAreErrors;
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