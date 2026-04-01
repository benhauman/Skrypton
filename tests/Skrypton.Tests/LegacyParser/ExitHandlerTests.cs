
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class ExitHandlerTests : TestBase
    {
        [TestMethod]
        public void DoNotCrashIfReachEndOfContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Exit", lineIndex1),
                new KeyWordToken("function", lineIndex1)
            };
            (new ExitHandler()).Process(tokens);
        }
    }
}
