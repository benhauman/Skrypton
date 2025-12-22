
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.ContentBreaking;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class StringBreakerTests : TestBase
    {
        [TestMethod, MyFact]
        public void VariableSetToStringContentIncludedQuotedContent()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("strValue = ", 0),
                    new StringToken("Test string with \"quoted\" content", 0),
                    new UnprocessedContentToken("\n", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "strValue = \"Test string with \"\"quoted\"\" content\"\n"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This tests the minimum escaped-content variable name that is possible (a blank variable name, escaped by square brackets)
        /// </summary>
        [TestMethod, MyFact]
        public void EmptyContentEscapedVariableNameIsSetToNumericValue()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(" = 1", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "[] = 1"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This tests the minimum escaped-content variable name that is possible (a blank variable name, escaped by square brackets)
        /// </summary>
        [TestMethod, MyFact]
        public void DeclaredEmptyContentEscapedVariableNameIsSetToNumericValue()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("Dim ", 0),
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(": ", 0),
                    new EscapedNameToken("[]", 0),
                    new UnprocessedContentToken(" = 1", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "Dim []: [] = 1"
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void InlineCommentsAreIdentifiedAsSuch()
        {
            // The StringBreaker will insert an EndOfStatementSameLineToken between the UnprocessedContentToken and InlineCommentToken
            // since that the later processes rely on end-of-statement tokens, even before an inline comment
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("WScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "WScript.Echo 1 ' Test"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// This recreates a bug where if there were line returns in the unprocessed content before what should be an inline comment, it
        /// wasn't realised that these were before the line that the comment should be inline with
        /// </summary>
        [TestMethod, MyFact]
        public void InlineCommentsAreIdentifiedAsSuchWhenAfterMultipleLinesOfContent()
        {
            // The StringBreaker will insert an EndOfStatementSameLineToken between the UnprocessedContentToken and InlineCommentToken
            // since that the later processes rely on end-of-statement tokens, even before an inline comment
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("\nWScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 1)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "\nWScript.Echo 1 ' Test"
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void REMCommentsAreIdentified()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new CommentToken(" Test", 0),
                    new UnprocessedContentToken("WScript.Echo 1", 1)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "REM Test\nWScript.Echo 1"
                ),
                new TokenSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void InlineREMCommentsAreIdentified()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("WScript.Echo 1", 0),
                    new EndOfStatementSameLineToken(0),
                    new InlineCommentToken(" Test", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "WScript.Echo 1 REM Test"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// If there were two comments on adjacent lines and the second has leading whitespace before the comment symbol then this whitespace would be incorrectly
        /// interpreted as unprocessed content, which must be terminated with an end-of-statement token. Instead, the content should be identified only as two
        /// comments.
        /// </summary>
        [TestMethod, MyFact]
        public void NonLineReturningWhiteSpaceBetweenCommentsIsIgnored()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new CommentToken(" Comment 1", 0),
                    new CommentToken(" Comment 2", 1)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "' Comment 1\n ' Comment 2"
                ),
                new TokenSetComparer()
            );
        }

        /// <summary>
        /// An end-of-statement token must be inserted between non-comment content and a comment - but there was a logic issue where this would be misidentified
        /// if the content before the comment was whitespace that was removed and a StringToken before that. This confirms the fix. (When a same-line end-of-
        /// statement token is inserted, the line index should not be incremented - this was included in the fix and is also demonstrated here).
        /// </summary>
        [TestMethod, MyFact]
        public void WhitespaceBetweenStringTokenAndCommentDoesNotPreventEndOfStatementBeingInserted()
        {
            myAssert.AreEqual(
                new IToken[]
                {
                    new UnprocessedContentToken("a = ", 0),
                    new StringToken("", 0),
                    new EndOfStatementSameLineToken(0),
                    new CommentToken(" Comment", 0)
                },
                StringBreaker.TestSegmentStringTest(TestCulture,
                    "a = \"\" ' Comment"
                ),
                new TokenSetComparer()
            );
        }
    }
}
