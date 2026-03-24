using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class TranslationResultExtensions
    {
        public static TranslationResult Add(this TranslationResult source, TranslatedStatement toAdd)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (toAdd == null)
                throw new ArgumentNullException(nameof(toAdd));

            return new TranslationResult(
                source.TranslatedStatements.Add(toAdd),
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, IReadOnlyCollection<TranslatedStatement> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (toAdd == null)
                throw new ArgumentNullException(nameof(toAdd));

            return new TranslationResult(
                source.TranslatedStatements.AddRange(toAdd),
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, TranslationResult toAdd)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (toAdd == null)
                throw new ArgumentNullException(nameof(toAdd));

            return new TranslationResult(
                source.TranslatedStatements.AddRange(toAdd.TranslatedStatements),
                source.ExplicitVariableDeclarations.AddRange(toAdd.ExplicitVariableDeclarations),
                source.UndeclaredVariablesAccessed.AddRange(toAdd.UndeclaredVariablesAccessed)
            );
        }

        public static TranslationResult AddExplicitVariableDeclarations(this TranslationResult source, IReadOnlyCollection<VariableDeclaration> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (toAdd == null)
                throw new ArgumentNullException(nameof(toAdd));

            return new TranslationResult(
                source.TranslatedStatements,
                source.ExplicitVariableDeclarations.AddRange(toAdd),
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult AddUndeclaredVariables(this TranslationResult source, IReadOnlyCollection<NameToken> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (toAdd == null)
                throw new ArgumentNullException(nameof(toAdd));

            return new TranslationResult(
                source.TranslatedStatements,
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed.AddRange(toAdd)
            );
        }
    }
}
