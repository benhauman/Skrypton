using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.CSharpWriter.CodeTranslation.StatementTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.CSharpWriter.Logging;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using System;
using System.Globalization;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    internal sealed class WithBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public WithBlockTranslator(
            CSharpName supportRefName,
            CSharpName envClassName,
            CSharpName envRefName,
            CSharpName outerClassName,
            CSharpName outerRefName,
            VBScriptNameRewriter nameRewriter,
            TempValueNameGenerator tempNameGenerator,
            ITranslateIndividualStatements statementTranslator,
            ITranslateValueSettingsStatements valueSettingStatementTranslator,
            ILogInformation logger)
            : base(supportRefName, envClassName, envRefName, outerClassName, outerRefName, nameRewriter, tempNameGenerator, statementTranslator, valueSettingStatementTranslator, logger)
        {
            _statementTranslator = statementTranslator ?? throw new ArgumentNullException(nameof(statementTranslator));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public TranslationResult Translate(WithBlock withBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
        {
            if (withBlock == null)
                throw new ArgumentNullException(nameof(withBlock));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            var translatedTargetReference = _statementTranslator.Translate(withBlock.Target, scopeAccessInformation, ExpressionReturnTypeOptions.Reference, _logger.Warning);
            var undeclaredVariables = translatedTargetReference.VariablesAccessed
                .Where(v => !scopeAccessInformation.IsDeclaredReference(v, _nameRewriter)).ToArray();
            foreach (var undeclaredVariable in undeclaredVariables)
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");

            var targetName = base._tempNameGenerator(new CSharpName("with"), scopeAccessInformation);
            var withBlockContentTranslationResult = Translate(
                withBlock.Content.ToNonNullImmutableList(),
                new ScopeAccessInformation(
                    withBlock,
                    scopeAccessInformation.ScopeDefiningParent,
                    scopeAccessInformation.ParentReturnValueNameIfAny,
                    scopeAccessInformation.ErrorRegistrationTokenIfAny,
                    new ScopeAccessInformation.DirectedWithReferenceDetails(
                        targetName,
                        withBlock.Target.Tokens.First().LineIndex
                    ),
                    scopeAccessInformation.ExternalDependencies,
                    scopeAccessInformation.Classes,
                    scopeAccessInformation.Functions,
                    scopeAccessInformation.Properties,
                    scopeAccessInformation.Constants,
                    scopeAccessInformation.Variables,
                scopeAccessInformation.StructureExitPoints
                ),
                indentationDepth
            );
            return new TranslationResult(
                withBlockContentTranslationResult.TranslatedStatements
                    .Insert(
                        new TranslatedStatement(TranslatedStatementKind.VariableDeclarationStatement,
                            string.Format(CultureInfo.InvariantCulture,
                                "var {0} = {1};",
                                targetName.Name,
                                translatedTargetReference.TranslatedContent
                            ),
                            indentationDepth,
                            withBlock.Target.Tokens.First().LineIndex
                        ),
                        0
                    ),
                withBlockContentTranslationResult.ExplicitVariableDeclarations,
                withBlockContentTranslationResult.UndeclaredVariablesAccessed.AddRange(undeclaredVariables)
            );
        }

        private TranslationResult Translate(NonNullImmutableList<ICodeBlock> blocks, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
        {
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            return base.TranslateCommon(
                base.GetWithinFunctionBlockTranslators(),
                blocks,
                scopeAccessInformation,
                indentationDepth
            );
        }
    }
}
