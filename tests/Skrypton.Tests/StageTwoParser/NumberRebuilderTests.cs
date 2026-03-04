
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.StageTwoParser.TokenCombining.NumberRebuilding;
using Skrypton.StageTwoParser.Tokens;
using Skrypton.Tests.Shared;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.StageTwoParser
{
    [TestClass]
    public class NumberRebuilderTests : TestBase
    {
        [TestMethod, MyFact]
        public void NegativeOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("-1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void BracketedNegativeOne()
        {
            myAssert.AreEqual(
                [
                    new OpenBrace(lineIndex1),
                    new NumericValueToken("-1", lineIndex1),
                    new CloseBrace(lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new OpenBrace(lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void PointOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken(".1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OnePointOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("1.1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void NegativeOnePointOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("-1.1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void NegativePointOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("-.1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new OperatorToken("-", lineIndex1),
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OnePlusNegativeOne()
        {
            myAssert.AreEqual(
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("+", lineIndex1),
                    new NumericValueToken("-1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
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
        public void NegativeOneAsNonBracketedArgument()
        {
            myAssert.AreEqual(
                [
                    new NameToken("fnc", lineIndex1),
                    new NumericValueToken("1.1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new NameToken("fnc", lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void PointOneAsNonBracketedArgument()
        {
            myAssert.AreEqual(
                [
                    new NameToken("fnc", lineIndex1),
                    new NumericValueToken(".1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new NameToken("fnc", lineIndex1),
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void ForLoopWithNegativeConstraints()
        {
            myAssert.AreEqual(
                [
                    new KeyWordToken("FOR", lineIndex1),
                    new NameToken("i", lineIndex1),
                    new ComparisonOperatorToken("=", lineIndex1),
                    new NumericValueToken("-1", lineIndex1),
                    new KeyWordToken("TO", lineIndex1),
                    new NumericValueToken("-4", lineIndex1),
                    new KeyWordToken("STEP", lineIndex1),
                    new NumericValueToken("-1", lineIndex1)
                ],
                NumberRebuilder.Rebuild(
                    [
                        new KeyWordToken("FOR", lineIndex1),
                        new NameToken("i", lineIndex1),
                        new ComparisonOperatorToken("=", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new KeyWordToken("TO", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("4", lineIndex1),
                        new KeyWordToken("STEP", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NumericValueToken("1", lineIndex1)
                    ]
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// When NameTokens are prefixed with a MemberAccessorOrDecimalPointToken, this is presumably because the content is wrapped in a "WITH" statement
        /// that will resolve the property / method access. As such, it shouldn't be assumed that a trailing dot is always a decimal point.
        /// </summary>
        [TestMethod, MyFact]
        public void DoNotTryToTreatMemberSeparatorRelyUponWithKeywordAsDecimalPoint()
        {
            myAssert.AreEqual(
                [
                    new MemberAccessorToken(hasLeadingWhiteSpace: false, lineIndex1),
                    new NameToken("Name", lineIndex1),
                ],
                NumberRebuilder.Rebuild(
                    [
                        new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                        new NameToken("Name", lineIndex1),
                    ]
                ),
                new TokenSetComparer()
            );
        }
    }
}
