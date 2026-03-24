using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter
{
    internal sealed class DefaultTranslatorOptions : ITranslatorOptions
    {
        private readonly Dictionary<string, bool> _suppressions = new Dictionary<string, bool>();

        public DefaultTranslatorOptions(IReadOnlyCollection<string> suppressions)
        {
            foreach (string suppression in suppressions.Select(x => x.ToUpperInvariant()).Distinct())
            {
                _suppressions.Add(suppression, true);
            }
        }

        public bool AcceptTranslationError(string errorKey) => _suppressions.TryGetValue(errorKey, out bool suppressed) && suppressed;
    }
}