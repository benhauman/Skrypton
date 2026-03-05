using System;
using Skrypton.CSharpWriter.CodeTranslation.StatementTranslation;
using Skrypton.LegacyParser.CodeBlocks.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class ITranslateIndividualStatementsExtensions
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
        /// </summary>
        public static TranslatedStatementContentDetails Translate(
            this ITranslateIndividualStatements statementTranslator,
            Statement statement,
            ScopeAccessInformation scopeAccessInformation,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException(nameof(statementTranslator));
            if (statement == null)
                throw new ArgumentNullException(nameof(statement));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (warningLogger == null)
                throw new ArgumentNullException(nameof(warningLogger));

            return Translate(statementTranslator, statement, scopeAccessInformation, ExpressionReturnTypeOptions.None, warningLogger);
        }

        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null codeExpression reference)
        /// </summary>
        public static TranslatedStatementContentDetails Translate(
            this ITranslateIndividualStatements statementTranslator,
            CodeExpression codeExpression,
            ScopeAccessInformation scopeAccessInformation,
            ExpressionReturnTypeOptions returnRequirements,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException(nameof(statementTranslator));
            if (codeExpression == null)
                throw new ArgumentNullException(nameof(codeExpression));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
                throw new ArgumentOutOfRangeException(nameof(returnRequirements));
            if (warningLogger == null)
                throw new ArgumentNullException(nameof(warningLogger));

            return Translate(statementTranslator, (Statement)codeExpression, scopeAccessInformation, returnRequirements, warningLogger);
        }

        private static TranslatedStatementContentDetails Translate(
            ITranslateIndividualStatements statementTranslator,
            Statement statement,
            ScopeAccessInformation scopeAccessInformation,
            ExpressionReturnTypeOptions returnRequirements,
            Action<string> warningLogger)
        {
            if (statementTranslator == null)
                throw new ArgumentNullException(nameof(statementTranslator));
            if (statement == null)
                throw new ArgumentNullException(nameof(statement));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
                throw new ArgumentOutOfRangeException(nameof(returnRequirements));
            if (warningLogger == null)
                throw new ArgumentNullException(nameof(warningLogger));

            return statementTranslator.Translate(
                statement.ToStageTwoParserExpression(scopeAccessInformation, returnRequirements, warningLogger),
                scopeAccessInformation,
                returnRequirements
            );
        }
    }
}
