
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.ContentBreaking;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.Shared.Comparers;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class TokenBreakerTests : TestBase
    {
        /// <summary>
        /// Previously, there was an error where a line break would result in a LineIndex increment for both the line break token and the token
        /// preceding it, rather than tokens AFTER the line break
        /// </summary>
        [TestMethod, MyFact]
        public void IncrementLineIndexAfterLineBreaks()
        {
            myAssert.AreEqual(
                [
                    new NameToken("Test", lineIndex1),
                    new NameToken("z", lineIndex1),
                    new EndOfStatementNewLineToken(lineIndex1)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("Test z\n", lineIndex1)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByWhitespace()
        {
            myAssert.AreEqual(
                [
                    new NameToken("a", lineIndex1),
                    new OperatorToken("&", lineIndex1),
                    new NameToken("b", lineIndex2)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a & _\nb", lineIndex1)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void UnderscoresAreLineContinuationsWhenTheyArePrecededByTokenBreakers()
        {
            myAssert.AreEqual(
                [
                    new NameToken("a", lineIndex1),
                    new OperatorToken("&", lineIndex1),
                    new NameToken("b", lineIndex2)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a&_\nb", lineIndex1)),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void DoNotConsiderUnderscoresToBeLineContinuationsWhenTheyArePartOfVariableNames()
        {
            myAssert.AreEqual(
                [
                    new NameToken("a_b", lineIndex1)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("a_b", lineIndex1)),
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
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("/", lineIndex1),
                    new NumericValueToken("0", lineIndex1)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1/0", lineIndex1)),
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
                [
                    new NumericValueToken("1", lineIndex1),
                    new OperatorToken("\\", lineIndex1),
                    new NumericValueToken("0", lineIndex1)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("1\\0", lineIndex1)),
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
                [
                    new NameToken("value", lineIndex1),
                    new ComparisonOperatorToken(OperatorKind.LessThan, "<", lineIndex1),
                    new ComparisonOperatorToken(OperatorKind.GreaterThan, ">", lineIndex1)
                ],
                TokenBreaker.BreakUnprocessedToken(new UnprocessedContentToken("value<>", lineIndex1)),
                new TokenSetComparer()
            );
        }
    }
}
