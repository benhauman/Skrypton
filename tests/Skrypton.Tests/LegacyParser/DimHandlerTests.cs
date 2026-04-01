
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class DimHandlerTests : TestBase
    {
        /// <summary>
        /// There was an issue where the argument separator tokens weren't being removed from DIM statements for multiple variables - this is the
        /// fail-before-fixing test for that issue
        /// </summary>
        [TestMethod]
        public void VariableSeparatorsAreCorrectlyRemovedAsProcessedContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Dim", lineIndex1),
                new NameToken("i", lineIndex1),
                new ArgumentSeparatorToken(lineIndex1),
                new NameToken("j", lineIndex1),
                new ArgumentSeparatorToken(lineIndex1),
                new NameToken("k", lineIndex1),
                new ArgumentSeparatorToken(lineIndex1),
                new NameToken("l", lineIndex1)
            };
            (new DimHandler()).Process(tokens);
            myAssert.AreEqual(0, tokens.Count);
        }
    }
}
