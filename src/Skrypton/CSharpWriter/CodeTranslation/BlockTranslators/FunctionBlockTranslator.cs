using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.CSharpWriter.CodeTranslation.StatementTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.CSharpWriter.Logging;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    public class FunctionBlockTranslator : CodeBlockTranslator
    {
        private readonly ITranslateIndividualStatements _statementTranslator;
        private readonly ILogInformation _logger;
        public FunctionBlockTranslator(
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

        public TranslationResult Translate(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation, int indentationDepth)
        {
            if (functionBlock == null)
                throw new ArgumentNullException(nameof(functionBlock));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            bool isSingleReturnValueStatementFunction = IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(functionBlock, scopeAccessInformation);
            CSharpName returnValueName = functionBlock.HasReturnValue
                ? _tempNameGenerator(new CSharpName($"{functionBlock.Name.Content}_retVal"), scopeAccessInformation.Extend(functionBlock, functionBlock.Statements.ToNonNullImmutableList())) // Ensure call Extend so that ScopeDefiningParent is the current function
                : null;
            TranslationResult translationResult = TranslationResult.Empty.Add(
                TranslateFunctionHeader(
                    functionBlock,
                    scopeAccessInformation,
                    returnValueName,
                    indentationDepth
                )
            );
            CSharpName errorRegistrationTokenIfAny;
            if (functionBlock.Statements.ToNonNullImmutableList().DoesScopeContainOnErrorResumeNext())
            {
                errorRegistrationTokenIfAny = _tempNameGenerator(new CSharpName("errOn"), scopeAccessInformation);
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(CultureInfo.InvariantCulture,
                        "var {0} = {1}.GETERRORTRAPPINGTOKEN();",
                        errorRegistrationTokenIfAny.Name,
                        _supportRefName.Name
                    ),
                    indentationDepth + 1,
                    functionBlock.Name.LineIndex
                ));
            }
            else
                errorRegistrationTokenIfAny = null;
            translationResult = translationResult.Add(
                Translate(
                    functionBlock.Statements.ToNonNullImmutableList(),
                    scopeAccessInformation.Extend(
                        functionBlock,
                        returnValueName,
                        errorRegistrationTokenIfAny,
                        functionBlock.Statements.ToNonNullImmutableList()
                    ),
                    isSingleReturnValueStatementFunction,
                    indentationDepth + 1
                )
            );
            int lineIndexForClosingScaffolding = translationResult.TranslatedStatements.Last().LineIndexOfStatementStartInSource;
            if (errorRegistrationTokenIfAny != null)
            {
                translationResult = translationResult.Add(new TranslatedStatement(
                    string.Format(CultureInfo.InvariantCulture,
                        "{0}.RELEASEERRORTRAPPINGTOKEN({1});",
                        _supportRefName.Name,
                        errorRegistrationTokenIfAny.Name
                    ),
                    indentationDepth + 1,
                    lineIndexForClosingScaffolding
                ));
            }
            if (functionBlock.HasReturnValue && !isSingleReturnValueStatementFunction)
            {
                // If this is an empty function then just render "return null" (TranslateFunctionHeader won't declare the return value reference)
                translationResult = translationResult
                    .Add(new TranslatedStatement(
                        string.Format(CultureInfo.InvariantCulture,
                            "return {0};",
                            functionBlock.Statements.Any() ? returnValueName.Name : "null"
                        ),
                        indentationDepth + 1,
                        lineIndexForClosingScaffolding
                    ));
            }
            return translationResult.Add(
                new TranslatedStatement("}", indentationDepth, lineIndexForClosingScaffolding)
            );
        }

        private TranslationResult Translate(
            NonNullImmutableList<ICodeBlock> blocks,
            ScopeAccessInformation scopeAccessInformation,
            bool isSingleReturnValueStatementFunction,
            int indentationDepth)
        {
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            NonNullImmutableList<BlockTranslationAttempter> blockTranslators;
            if (isSingleReturnValueStatementFunction)
            {
                blockTranslators = new NonNullImmutableList<BlockTranslationAttempter>()
                    .Add(TryToTranslateValueSettingStatementAsSimpleFunctionValueReturner)
                    .Add(TryToTranslateBlankLine)
                    .Add(TryToTranslateComment);
            }
            else
                blockTranslators = base.GetWithinFunctionBlockTranslators();

            return base.TranslateCommon(
                blockTranslators,
                blocks,
                scopeAccessInformation,
                indentationDepth
            );
        }

        private TranslationResult TryToTranslateValueSettingStatementAsSimpleFunctionValueReturner(
            TranslationResult translationResult,
            ICodeBlock block,
            ScopeAccessInformation scopeAccessInformation,
            int indentationDepth)
        {
            if (translationResult == null)
                throw new ArgumentNullException(nameof(translationResult));
            if (block == null)
                throw new ArgumentNullException(nameof(block));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            ValueSettingStatement valueSettingStatement = block as ValueSettingStatement;
            if (valueSettingStatement == null)
                return null;

            TranslatedStatementContentDetails translatedStatementContentDetails = _statementTranslator.Translate(
                valueSettingStatement.Expression,
                scopeAccessInformation,
                (valueSettingStatement.ValueSetType == ValueSetTypeOptions.Set)
                    ? ExpressionReturnTypeOptions.Reference
                    : ExpressionReturnTypeOptions.Value,
                _logger.Warning
            );
            NameToken[] undeclaredVariables = translatedStatementContentDetails.VariablesAccessed.Where(v => !scopeAccessInformation.IsDeclaredReference(v, _nameRewriter)).ToArray();
            foreach (NameToken undeclaredVariable in undeclaredVariables)
            {
                _logger.Warning("Undeclared variable: \"" + undeclaredVariable.Content + "\" (line " + (undeclaredVariable.LineIndex + 1) + ")");
            }

            return translationResult
                .Add(new TranslatedStatement(
                    "return " + translatedStatementContentDetails.TranslatedContent + ";",
                    indentationDepth,
                    valueSettingStatement.Expression.Tokens.First().LineIndex
                ))
                .AddUndeclaredVariables(undeclaredVariables);
        }

        private TranslatedStatement[] TranslateFunctionHeader(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation, CSharpName returnValueNameIfAny, int indentationDepth)
        {
            if (functionBlock == null)
                throw new ArgumentNullException(nameof(functionBlock));
            if (functionBlock.HasReturnValue && (returnValueNameIfAny == null))
                throw new ArgumentException("returnValueNameIfAny must not be null if functionBlock.HasReturnValue is true");
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");

            StringBuilder content = new StringBuilder();
            content.Append(functionBlock.IsPublic ? "public" : "private");
            content.Append(' ');
            content.Append(functionBlock.HasReturnValue ? "object" : "void");  // #pragma warning disable CS0219
            content.Append(' ');
            content.Append(_nameRewriter.GetMemberAccessTokenName(functionBlock.Name));
            content.Append('(');
            int numberOfParameters = functionBlock.Parameters.Count();
            for (int index = 0; index < numberOfParameters; index++)
            {
                AbstractFunctionBlock.Parameter parameter = functionBlock.Parameters.ElementAt(index);
                if (parameter.ByRef)
                    content.Append("ref ");
                content.Append("object ");
                content.Append(_nameRewriter.GetMemberAccessTokenName(parameter.Name));
                if (index < (numberOfParameters - 1))
                    content.Append(", ");
            }
            content.Append(')');

            List<TranslatedStatement> translatedStatements = new List<TranslatedStatement>();
            if (functionBlock.IsDefault)
                translatedStatements.Add(new TranslatedStatement("[" + typeof(IsDefaultAttribute).FullName + "]", indentationDepth, functionBlock.Name.LineIndex));
            PropertyBlock property = functionBlock as PropertyBlock;
            if (property != null)
            {
                // All property blocks that are translated into C# methods needs to be decorated with the [TranslatedProperty] attribute. The [TranslatedProperty] attribute
                // was originally intended only for indexed properties (which C# can only support one of per class but VBScript classes can have as many as they like) but
                // a class with an indexed property will be emitted to inherit from TranslatedPropertyIReflectImplementation, which will try to identify properties based
                // upon the presence of [TranslatedProperty] attributes - if some (ie. indexed properties) have these and others (non-indexed properties) don't then it
                // will result in runtime failures. So we could apply the attribute to indexed properties and all properties within classes that have at least one
                // indexed property but that feels like complications for little benefit so I think it's easier to just put it on ALL from-property methods.
                translatedStatements.Add(
                    new TranslatedStatement(
                        string.Format(CultureInfo.InvariantCulture,
                            "[TranslatedProperty({0})]", // Note: Safe to assume that using statements are present for the namespace that contains TranslatedProperty
                            property.Name.Content.ToLiteral()
                        ),
                        indentationDepth,
                        functionBlock.Name.LineIndex
                    )
                );
            }
            translatedStatements.Add(new TranslatedStatement(content.ToString(), indentationDepth, functionBlock.Name.LineIndex));
            translatedStatements.Add(new TranslatedStatement("{", indentationDepth, functionBlock.Name.LineIndex));
            if (functionBlock.HasReturnValue && functionBlock.Statements.Any() && !IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(functionBlock, scopeAccessInformation))
            {
                translatedStatements.Add(new TranslatedStatement(
                    base.TranslateVariableInitialization(
                        new VariableDeclaration(
                            new DoNotRenameNameToken(
                                returnValueNameIfAny.Name.ToUpperX(),
                                functionBlock.Name.LineIndex
                            ),
                            VariableDeclarationScopeOptions.Private,
                            null // Not declared as an array
                        ),
                        ScopeLocationOptions.WithinFunctionOrPropertyOrWith,
                        asUnreferencedVar: false,
                        indentationDepth + 1
                    ),
                    indentationDepth + 1,
                    functionBlock.Name.LineIndex
                ));
            }
            return translatedStatements.ToArray();
        }

        /// <summary>
        /// If a function or property only contains a single executable block, which is a return statement, then this can be translated into a simple return
        /// statement in the C# output (as opposed to having to maintain a temporary variable for the return value in case there are various manipulations
        /// of it or error-handling or any other VBScript oddnes required)
        /// </summary>
        private bool IsSingleReturnValueStatementFunctionWithoutAnyByRefMappings(AbstractFunctionBlock functionBlock, ScopeAccessInformation scopeAccessInformation)
        {
            if (functionBlock == null)
                throw new ArgumentNullException(nameof(functionBlock));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));

            ICodeBlock[] executableStatements = functionBlock.Statements.Where(s => !(s is INonExecutableCodeBlock)).ToArray();
            if (executableStatements.Length != 1)
                return false;

            ValueSettingStatement valueSettingStatement = executableStatements.Single() as ValueSettingStatement;
            if (valueSettingStatement == null)
                return false;

            if (valueSettingStatement.ValueToSet.Tokens.Count() != 1)
                return false;

            NameToken valueToSetTokenAsNameToken = valueSettingStatement.ValueToSet.Tokens.Single() as NameToken;
            if (valueToSetTokenAsNameToken == null)
                return false;

            if (_nameRewriter.GetMemberAccessTokenName(valueToSetTokenAsNameToken) != _nameRewriter.GetMemberAccessTokenName(functionBlock.Name))
                return false;

            // If there is no return value (ie. it's a SUB or a LET/SET PROPERTY accessor) then this can't apply (not only can this simple single-line
            // return format not be used but a runtime error is required if the value-setting statement targets the name of a SUB)
            if (!functionBlock.HasReturnValue)
                return false;

            // If any values need aliasing in order to perform this "one liner" then it won't be possible to represent it a simple one-line return, it will
            // need a try..finally setting up to create the alias(es), use where required and then map the values back over the original(s).
            scopeAccessInformation = scopeAccessInformation.Extend(functionBlock, functionBlock.Statements.ToNonNullImmutableList());
            FuncByRefArgumentMapper byRefArgumentMapper = new FuncByRefArgumentMapper(_nameRewriter, _tempNameGenerator, _logger);
            NonNullImmutableList<FuncByRefMapping> byRefArgumentsToMap = byRefArgumentMapper.GetByRefArgumentsThatNeedRewriting(
                valueSettingStatement.Expression.ToStageTwoParserExpression(scopeAccessInformation, ExpressionReturnTypeOptions.NotSpecified, _logger.Warning),
                scopeAccessInformation,
                new NonNullImmutableList<FuncByRefMapping>()
            );
            if (byRefArgumentsToMap.Any())
                return false;

            return !valueSettingStatement.Expression.Tokens.Any(
                t => (t is NameToken) && (_nameRewriter.GetMemberAccessTokenName(t) == _nameRewriter.GetMemberAccessTokenName(functionBlock.Name))
            );
        }
    }
}
