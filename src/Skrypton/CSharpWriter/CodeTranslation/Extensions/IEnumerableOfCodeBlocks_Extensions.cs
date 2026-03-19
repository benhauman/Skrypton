using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class IEnumerableOfCodeBlocksExtensions
    {
        public static bool DoesScopeContainOnErrorResumeNext(this IEnumerable<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));

            foreach (var block in blocks)
            {
                if (block == null)
                    throw new ArgumentException("Null reference encountered in blocks set");

                if (block is OnErrorResumeNext)
                    return true;

                if (block is IDefineScope)
                {
                    // If this block defines its own scope then it doesn't matter if it contains an OnErrorResumeNext statement
                    // within it, it won't affect this scope
                    continue;
                }

                var blockWithNestedContent = block as IHaveNestedContent;
                if ((blockWithNestedContent != null) && blockWithNestedContent.AllExecutableBlocks.ToNonNullImmutableList().DoesScopeContainOnErrorResumeNext())
                    return true;
            }
            return false;
        }

        public static IEnumerable<IToken> EnumerateAllTokens(this IEnumerable<ICodeBlock> blocks)
        {
            if (blocks == null)
                throw new ArgumentNullException(nameof(blocks));

            foreach (var block in blocks)
            {
                if (block == null)
                {
                    throw new ArgumentException("Null reference encountered in blocks set");
                }

                var nonNestedExpressionContainingBlock = block as IHaveNonNestedExpressions;
                IEnumerable<Statement> expressionsToInterrogate;
                if (nonNestedExpressionContainingBlock != null)
                {
                    expressionsToInterrogate = nonNestedExpressionContainingBlock.NonNestedExpressions ?? [];
                }
                else if (block is Statement statement)
                {
                    expressionsToInterrogate = new[] { statement };
                }
                else
                {
                    expressionsToInterrogate = [];
                }
                foreach (var token in expressionsToInterrogate.SelectMany(e => e.Tokens))
                {
                    yield return token;
                }

                var nestedContentBlock = block as IHaveNestedContent;
                if (nestedContentBlock != null)
                {
                    foreach (var nestedToken in EnumerateAllTokens(nestedContentBlock.AllExecutableBlocks))
                    {
                        yield return nestedToken;
                    }
                }
            }
        }
    }
}
