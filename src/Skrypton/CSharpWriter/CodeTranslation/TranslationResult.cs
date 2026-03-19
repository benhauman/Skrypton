using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.Tokens.Basic;
using System;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    public sealed class TranslationResult
    {
        public TranslationResult(
            NonNullImmutableList<TranslatedStatement> translatedStatements,
            NonNullImmutableList<VariableDeclaration> explicitVariableDeclarations,
            NonNullImmutableList<NameToken> undeclaredVariablesAccessed)
        {
            TranslatedStatements = translatedStatements ?? throw new ArgumentNullException(nameof(translatedStatements));
            ExplicitVariableDeclarations = explicitVariableDeclarations ?? throw new ArgumentNullException(nameof(explicitVariableDeclarations));
            UndeclaredVariablesAccessed = undeclaredVariablesAccessed ?? throw new ArgumentNullException(nameof(undeclaredVariablesAccessed));
        }


        public static TranslationResult Empty
        {
            get
            {
                return new TranslationResult(
                    new NonNullImmutableList<TranslatedStatement>(),
                    new NonNullImmutableList<VariableDeclaration>(),
                    new NonNullImmutableList<NameToken>()
                );
            }
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<TranslatedStatement> TranslatedStatements { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<VariableDeclaration> ExplicitVariableDeclarations { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NonNullImmutableList<NameToken> UndeclaredVariablesAccessed { get; private set; }
    }
}
