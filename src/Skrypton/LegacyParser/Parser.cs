using System;
using System.Collections.Generic;
using System.Globalization;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.ContentBreaking;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.LegacyParser
{
    public static class Parser
    {
        public static CSharpWriter.CodeTranslation.IOutermostScope ParseToOutermostScope(CultureInfo culture, string scriptContent)
        {
            //var rootName = new CSharpWriter.CodeTranslation.Extensions.DoNotRenameNameToken("Root", 0);
            CSharpWriter.CodeTranslation.CSharpName rootName = new CSharpWriter.CodeTranslation.CSharpName("Root");
            IEnumerable<ICodeBlock> rootCodeBlocks = Parser.Parse(culture, scriptContent);
            //CodeBlockCollection rootStatements = new CodeBlockCollection(rootCodeBlocks);
            return new CSharpWriter.CodeTranslation.OutermostScope(rootName, new CSharpWriter.Lists.NonNullImmutableList<ICodeBlock>(rootCodeBlocks));
        }
        public static IReadOnlyCollection<ICodeBlock> Parse(CultureInfo culture, string scriptContent)
        {
            if (string.IsNullOrWhiteSpace(scriptContent))
                throw new ArgumentException("Null/blank scriptContent specified");

            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(culture,
                scriptContent.Replace("\r\n", "\n")
            );

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            var atomTokens = new List<IToken>();
            //foreach (var token in tokens)
            for (int ix = 0; ix < tokens.Count; ix++)
            {
                IToken token = tokens[ix];
                if (token is UnprocessedContentToken uct)
                {
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken(uct));
                }
                else
                {
                    atomTokens.Add(token);
                }
            }

            // Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
            return (CodeBlockHandler.RootBlock).Process(atomTokens, out var _);
        }
    }
}
