
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser.CodeBlocks.Handlers;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
//#using Xunit#;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class ExitHandlerTests
    {
        [TestMethod, MyFact]
        public void DoNotCrashIfReachEndOfContent()
        {
            var tokens = new List<IToken>
            {
                new KeyWordToken("Exit", 0),
                new KeyWordToken("function", 0)
            };
            (new ExitHandler()).Process(tokens);
        }
    }
}
