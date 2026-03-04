
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.StageTwoParser.TokenCombining.OperatorCombinations;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.StageTwoParser
{
    [TestClass]
    public class OperatorCombinerTests : TestBase
    {
        [TestMethod, MyFact]
        public void OnePlusNegativeOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("-", lineIndex1),
                    new NumericValueToken("1", lineIndex1)
                ],
                OperatorCombiner.Combine(
                    [
                        new NumericValueToken("1", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OneMinusNegativeOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("+", lineIndex1),
                    new NumericValueToken("1", lineIndex1)
                ],
                OperatorCombiner.Combine(
                    [
                        new NumericValueToken("1", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
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
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("*", lineIndex1),
                    new BuiltInFunctionToken("CInt", lineIndex1),
                    new OpenBrace(lineIndex1),
                    new NumericValueToken("1", lineIndex1),
                    new CloseBrace(lineIndex1)
                ],
                OperatorCombiner.Combine(
                    [
                        new NumericValueToken("1", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                TokenSetComparer.Instance
            );
        }

        [TestMethod, MyFact]
        public void TwoGreaterThanOrEqualToOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("2", lineIndex1),
                    new ComparisonOperatorToken(">=", lineIndex1),
                    new NumericValueToken("1", lineIndex1)
                ],
                OperatorCombiner.Combine(
                    [
                        new NumericValueToken("2", lineIndex1),
                        new ComparisonOperatorToken(">", lineIndex1),
                        new ComparisonOperatorToken("=", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }
    }
}
