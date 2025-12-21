
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class DimHandlerTests
    {
        /// <summary>
        /// There was an issue where the argument separator tokens weren't being removed from DIM statements for multiple variables - this is the
        /// fail-before-fixing test for that issue
        /// </summary>
        [TestMethod, MyFact]
        public void VariableSeparatorsAreCorrectlyRemovedAsProcessedContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Dim", 0),
                new NameToken("i", 0),
                new ArgumentSeparatorToken(0),
                new NameToken("j", 0),
                new ArgumentSeparatorToken(0),
                new NameToken("k", 0),
                new ArgumentSeparatorToken(0),
                new NameToken("l", 0)
            };
            (new DimHandler()).Process(tokens);
            myAssert.AreEqual(0, tokens.Count);
        }
    }
}
