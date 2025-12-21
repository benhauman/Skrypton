
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.ContentBreaking;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class TokenBreakerTests
    {
        /// <summary>
        /// Previously, there was an error where a line break would result in a LineIndex increment for both the line break token and the token
        /// preceding it, rather than tokens AFTER the line break
        /// </summary>
        [TestMethod, MyFact]
        public void IncrementLineIndexAfterLineBreaks()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("Test", 0),
                    new NameToken("z", 0),
                    new EndOfStatementNewLineToken(0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("Test z\n", 0)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByWhitespace()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("a", 0),
                    new OperatorToken("&", 0),
                    new NameToken("b", 1)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a & _\nb", 0)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByTokenBreakers()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("a", 0),
                    new OperatorToken("&", 0),
                    new NameToken("b", 1)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a&_\nb", 0)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void DoNotConsiderUnderscoresToBeLineContinuationsWhenTheyArePartOfVariableNames()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("a_b", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a_b", 0)),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// I realised that "1/0" wasn't being correctly broken down since the "/" wasn't being considered a "Token Break Character" and so the "1/0" was being
        /// interpreted as a NameToken, instead of two numeric value tokens and an operator.
        /// </summary>
        [TestMethod, MyFact]
        public void EnsureThatDivisionOperatorsAreRecognised()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("/", 0),
                    new NumericValueToken("0", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1/0", 0)),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This is the same issue as that for which the EnsureThatDivisionOperatorsAreRecognised test was added, but for the integer division opereator (back
        /// slash, rather than forward)
        /// </summary>
        [TestMethod, MyFact]
        public void EnsureThatIntegerDivisionOperatorsAreRecognised()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NumericValueToken("1", 0),
                    new OperatorToken("\\", 0),
                    new NumericValueToken("0", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1\\0", 0)),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This is an issue identified in testing real content. The first of the following was correctly parsed while the second wasn't -
        ///   value <> ""
        ///   value<> ""
        /// It should be broken down into four tokens:
        ///   NameToken:"value"
        ///   ComparisonOperationToken:"<"
        ///   ComparisonOperationToken:">"
        ///   StringToken:""
        /// The TokenBreaker would then get an UnprocessedContentToken with content "value<>" which it needs to break into three.
        /// </summary>
        [TestMethod, MyFact]
        public void LessThanComparisonOperatorIndicatesTokenBreakRegardlessOfWhitespace()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new NameToken("value", 0),
                    new ComparisonOperatorToken("<", 0),
                    new ComparisonOperatorToken(">", 0)
                },
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("value<>", 0)),
                new TokenSetComparer()
            );
        }
    }
}
