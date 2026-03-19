using Skrypton.StageTwoParser.ExpressionParsing;
using System.Collections.Generic;

namespace Skrypton.CSharpWriter.CodeTranslation.StatementTranslation
{
    public interface ITranslateIndividualStatements
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null parsingExpression reference)
        /// </summary>
        TranslatedStatementContentDetails TranslateParsingExpression(ParsingExpression parsingExpression, ScopeAccessInformation scopeAccessInformation, ExpressionReturnTypeOptions returnRequirements);

        /// <summary>
        /// This generates the content that initialises a new IProvideCallArguments instance, based upon the specified argument values. This will throw
        /// an exception for null arguments or an argumentValues set containing any null references. It will never return null, it will raise an exception
        /// if unable to satisfy the request.
        /// </summary>
        TranslatedStatementContentDetails TranslateAsArgumentProvider(
            IEnumerable<ParsingExpression> argumentValues,
            ScopeAccessInformation scopeAccessInformation,
            bool forceAllArgumentsToBeByVal
        );
    }
}
