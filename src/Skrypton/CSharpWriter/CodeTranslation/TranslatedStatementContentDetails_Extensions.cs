using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    internal static class TranslatedStatementContentDetailsExtensions
    {
        /// <summary>
        /// This will never be null
        /// </summary>
        public static IReadOnlyCollection<NameToken> GetUndeclaredVariablesAccessed(
            this TranslatedStatementContentDetails source,
            ScopeAccessInformation scopeAccessInformation,
            VBScriptNameRewriter nameRewriter)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (nameRewriter == null)
                throw new ArgumentNullException(nameof(nameRewriter));

            return source.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(v, nameRewriter))
                .ToNonNullImmutableList();
        }
    }
}
