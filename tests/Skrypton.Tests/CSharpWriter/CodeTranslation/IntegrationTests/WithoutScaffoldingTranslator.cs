using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using System;
using System.Globalization;
using System.Linq;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public static class WithoutScaffoldingTranslator
    {
        public static NonNullImmutableList<string> DefaultConsoleExternalDependencies = new NonNullImmutableList<string>().Add("wscript");

        /// <summary>
        /// This will never return null or an array containing any nulls, blank values or values with leading or trailing whitespace or values containing line
        /// returns (this format makes the myAssert.AreEquals easier, where it can make array comparisons easily but not any IEnumerable implementation)
        /// </summary>
        public static string[] GetTranslatedStatements(CultureInfo culture, string content, NonNullImmutableList<string> externalDependencies)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (externalDependencies == null)
                throw new ArgumentNullException("externalDependencies");

            return DefaultTranslator
                .Translate(culture,
                    content,
                    externalDependencies,
                    OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding,
                    renderCommentsAboutUndeclaredVariables: false
                )
                .Select(s => s.Content)
                .Where(s => s != "")
                .ToArray();
        }
    }
}
