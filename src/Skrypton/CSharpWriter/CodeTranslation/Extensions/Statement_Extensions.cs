using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.StageTwoParser.ExpressionParsing;
using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class StatementExtensions
    {
        /// <summary>
        /// This will never return null, it will raise an exception if unable to satisfy the request (this includes the case of a null statement reference)
        /// </summary>
        public static ParsingExpression ToStageTwoParserExpression(
            this Statement statement,
            ScopeAccessInformation scopeAccessInformation,
            ExpressionReturnTypeOptions returnRequirements,
            Action<string> warningLogger)
        {
            if (statement == null)
                throw new ArgumentNullException(nameof(statement));
            if (scopeAccessInformation == null)
                throw new ArgumentNullException(nameof(scopeAccessInformation));
            if (!Enum.IsDefined(typeof(ExpressionReturnTypeOptions), returnRequirements))
                throw new ArgumentOutOfRangeException(nameof(returnRequirements));
            if (warningLogger == null)
                throw new ArgumentNullException(nameof(warningLogger));

            // The BracketStandardisedTokens property should only be used if this is a non-value-returning statement (eg. "Test" or "Test 1"
            // or "Test(a)", which would be translated into "Test()", "Test(1)" or "Test((a))", respectively) since that is the only time
            // that brackets appear "optional". When this statement's return value is considered (eg. the "Test(1)" in "a = Test(1)"), the
            // brackets will already be in a format in valid VBScript that matches what would be expected in C#.
            IReadOnlyCollection<IToken> xTokens = (returnRequirements == ExpressionReturnTypeOptions.None) ? statement.GetBracketStandardisedTokens(scopeAccessInformation.DirectedWithReferenceIfAny?.AsToken()) : statement.Tokens.ToArray();
            ParsingExpression[] expressions = ExpressionGenerator.GenerateExpressions(1, 1, xTokens, WithStatementInformation.TryGet(scopeAccessInformation), warningLogger);
            if (expressions.Length != 1)
                throw new ArgumentException("Statement translation should always result in a single codeExpression being generated");
            return expressions[0];
        }
    }
}
