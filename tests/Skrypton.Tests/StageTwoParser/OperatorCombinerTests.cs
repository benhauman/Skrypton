
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.StageTwoParser.TokenCombining.OperatorCombinations;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.StageTwoParser
{
    [TestClass]
    public class OperatorCombinerTests
    {
        [TestMethod, MyFact]
        public void OnePlusNegativeOne()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("-", 0),
                    new NumericValueToken("1", 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new OperatorToken("+", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OneMinusNegativeOne()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("+", 0),
                    new NumericValueToken("1", 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new OperatorToken("-", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OneMultipliedByPlusOne()
        {
            // When operators are removed entirely by the OperatorCombiner, if they are removed from in front of numeric values, the numeric value is wrapped
            // up in a CInt, CLng or CDbl call so that it is clear to the processing following it that it is not a numeric literal (but a function is chosen
            // that will its value - so here, for the small value 1 it is CInt).
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("*", 0),
                    new BuiltInFunctionToken("CInt", 0),
                    new OpenBrace(0),
                    new NumericValueToken("1", 0),
                    new CloseBrace(0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new OperatorToken("*", 0),
                        new OperatorToken("+", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                TokenSetComparer.Instance
            );
        }

        [TestMethod, MyFact]
        public void TwoGreaterThanOrEqualToOne()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("2",0),
                    new ComparisonOperatorToken(">=", 0),
                    new NumericValueToken("1", 0)
                },
                OperatorCombiner.Combine(
                    new IToken[]
                    {
                        new NumericValueToken("2",0),
                        new ComparisonOperatorToken(">", 0),
                        new ComparisonOperatorToken("=", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
