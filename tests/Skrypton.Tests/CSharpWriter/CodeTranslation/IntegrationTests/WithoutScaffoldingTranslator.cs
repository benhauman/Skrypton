using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using System;
using System.Globalization;
using System.Linq;
using Skrypton.CSharpWriter.CodeTranslation;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public static class WithoutScaffoldingTranslator // use 'DefaultCSharpTranslation'
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
                    OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding, // Executable:159 tests
                    renderCommentsAboutUndeclaredVariables: false
                )
                .Select(s => s.Content)
                .Where(s => s != "") // 129 tests
                .ToArray();
        }
    }

    internal static class DefaultCSharpTranslation
    {
        internal static string[] GetTranslatedStatements(CultureInfo culture, string vbsSource, NonNullImmutableList<string> externalDependencies)
        {
            string[] output = Skrypton.CSharpWriter.DefaultTranslator.Translate(culture, vbsSource, externalDependencies, OuterScopeBlockTranslator.OutputTypeOptions.Executable, true)
                .Select(s => RenderTranslatedStatement(s))
                .ToArray();
            return output; // later: string.Join("\r\n", output)
        }
        private static string RenderTranslatedStatement(TranslatedStatement s)
        {
            if (s.IndentationDepth == 0)
                return s.Content;
            if (!s.HasContent)
                return s.Content; // no indention for blank lines
            string txt = new string(' ', s.IndentationDepth * 4) + s.Content;
            return txt;
        }
    }
}
