using System;
using System.Collections.Generic;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    public static class TranslationResult_Extensions
    {
        public static TranslationResult Add(this TranslationResult source, TranslatedStatement toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements.Add(toAdd),
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, IEnumerable<TranslatedStatement> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements.AddRange(toAdd),
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult Add(this TranslationResult source, TranslationResult toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements.AddRange(toAdd.TranslatedStatements),
                source.ExplicitVariableDeclarations.AddRange(toAdd.ExplicitVariableDeclarations),
                source.UndeclaredVariablesAccessed.AddRange(toAdd.UndeclaredVariablesAccessed)
            );
        }

        public static TranslationResult AddExplicitVariableDeclarations(this TranslationResult source, IEnumerable<VariableDeclaration> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements,
                source.ExplicitVariableDeclarations.AddRange(toAdd.ToNonNullImmutableList()),
                source.UndeclaredVariablesAccessed
            );
        }

        public static TranslationResult AddUndeclaredVariables(this TranslationResult source, IEnumerable<NameToken> toAdd)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (toAdd == null)
                throw new ArgumentNullException("toAdd");

            return new TranslationResult(
                source.TranslatedStatements,
                source.ExplicitVariableDeclarations,
                source.UndeclaredVariablesAccessed.AddRange(toAdd.ToNonNullImmutableList())
            );
        }
    }
}
