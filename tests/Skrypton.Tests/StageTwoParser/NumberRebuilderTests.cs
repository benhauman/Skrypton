
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
    public class NumberRebuilderTests
    {
        [TestMethod, MyFact]
        public void NegativeOne()
        {
            myAssert.AreEqual(
                new[]
                {
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void BracketedNegativeOne()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new OpenBrace(0),
                    new NumericValueToken("-1", 0),
                    new CloseBrace(0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OpenBrace(0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new CloseBrace(0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void PointOne()
        {
            myAssert.AreEqual(
                new[]
                {
                    new NumericValueToken(".1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OnePointOne()
        {
            myAssert.AreEqual(
                new[]
                {
                    new NumericValueToken("1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void NegativeOnePointOne()
        {
            myAssert.AreEqual(
                new[]
                {
                    new NumericValueToken("-1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void NegativePointOne()
        {
            myAssert.AreEqual(
                new[]
                {
                    new NumericValueToken("-.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new OperatorToken("-", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void OnePlusNegativeOne()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("+", 0),
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
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
        public void NegativeOneAsNonBracketedArgument()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("fnc", 0),
                    new NumericValueToken("1.1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NameToken("fnc", 0),
                        new NumericValueToken("1", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void PointOneAsNonBracketedArgument()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("fnc", 0),
                    new NumericValueToken(".1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new NameToken("fnc", 0),
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NumericValueToken("1", 0)
                    }
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void ForLoopWithNegativeConstraints()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new KeyWordToken("FOR", 0),
                    new NameToken("i", 0),
                    new ComparisonOperatorToken("=", 0),
                    new NumericValueToken("-1", 0),
                    new KeyWordToken("TO", 0),
                    new NumericValueToken("-4", 0),
                    new KeyWordToken("STEP", 0),
                    new NumericValueToken("-1", 0)
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new KeyWordToken("FOR", 0),
                        new NameToken("i", 0),
                        new ComparisonOperatorToken("=", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0),
                        new KeyWordToken("TO", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("4", 0),
                        new KeyWordToken("STEP", 0),
                        new OperatorToken("-", 0),
                        new NumericValueToken("1", 0)
                    }
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
                new IToken[]
                {
                    new MemberAccessorToken(0),
                    new NameToken("Name", 0),
                },
                NumberRebuilder.Rebuild(
                    new IToken[]
                    {
                        new MemberAccessorOrDecimalPointToken(".", 0),
                        new NameToken("Name", 0),
                    }
                ),
                new TokenSetComparer()
            );
        }
    }
}
