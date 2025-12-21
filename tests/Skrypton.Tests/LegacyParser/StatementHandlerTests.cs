
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class StatementHandlerTests
    {
        /// <summary>
        /// Only the first non-bracketed equality sign in a statement may indicate the separation between the value-to-set and the expression-to-set-it-to
        /// in a value-setting-statement, any subsequent equals signs are comparison operators (C# uses "==" for comparisons, as opposed to "=" for setting
        /// values, which is clearer.. but this is VBScript)
        /// </summary>
        [TestMethod, MyFact]
        public void SubsequentEqualsTokensInValueSettingStatementAreComparisonOperators()
        {
            var statement = (new StatementHandler()).Process(new List<IToken>
            {
                new NameToken("bMatch", 0),
                new ComparisonOperatorToken("=", 0),
                new NumericValueToken("1", 0),
                new ComparisonOperatorToken("=", 0),
                new NumericValueToken("2", 0)
            });

            myAssert.IsType<ValueSettingStatement>(statement);

            var valueSettingStatement = (ValueSettingStatement)statement;
            myAssert.AreEqual(ValueSetTypeOptions.Let, valueSettingStatement.ValueSetType);
            myAssert.AreEqual(
                new IToken[] { new NameToken("bMatch", 0) },
                valueSettingStatement.ValueToSet.Tokens,
                new TokenSetComparer()
            );
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new ComparisonOperatorToken("=", 0),
                    new NumericValueToken("2", 0)
                },
                valueSettingStatement.Expression.Tokens,
                new TokenSetComparer()
            );
        }
    }
}
