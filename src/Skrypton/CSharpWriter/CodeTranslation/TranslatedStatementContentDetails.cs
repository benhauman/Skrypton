using System;
using System.Collections.Generic;
using System.Diagnostics;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("({VariablesAccessed.Count}){TranslatedContent}")]
    public class TranslatedStatementContentDetails // base class of 'TranslatedStatementContentDetailsWithContentType'
    {
        public TranslatedStatementContentDetails(string translatedContent, IReadOnlyCollection<NameToken> variablesAccessed)
        {
            if (string.IsNullOrWhiteSpace(translatedContent))
                throw new ArgumentException("Null/blank translatedContent specified");
            TranslatedContent = translatedContent;
            VariablesAccessed = variablesAccessed ?? throw new ArgumentNullException(nameof(variablesAccessed));
        }

        /// <summary>
        /// This will never return null or blank
        /// </summary>
        public string TranslatedContent { get; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public IReadOnlyCollection<NameToken> VariablesAccessed { get; private set; }
    }
}
