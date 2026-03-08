using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using System;
using System.Globalization;
using System.Linq;
using Skrypton.CSharpWriter.CodeTranslation;
using System.Text;
using Microsoft.Testing.Extensions.VSTestBridge;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    public static class WithoutScaffoldingTranslator // use 'DefaultCSharpTranslation'
    {
        public static NonNullImmutableList<string> DefaultConsoleExternalDependencies = new NonNullImmutableList<string>().Add("WScript");

        /// <summary>
        /// This will never return null or an array containing any nulls, blank values or values with leading or trailing whitespace or values containing line
        /// returns (this format makes the myAssert.AreEquals easier, where it can make array comparisons easily but not any IEnumerable implementation)
        /// </summary>
        public static string[] GetTranslatedStatements(CultureInfo culture, string content, NonNullImmutableList<string> externalDependencies)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (externalDependencies == null)
                throw new ArgumentNullException(nameof(externalDependencies));

            return DefaultTranslator.TranslateWithoutScaffolding(culture, content, externalDependencies) // Executable:159 tests
                .Select(s => s.Content)
                .Where(s => s != "") // 129 tests
                .ToArray();
        }
    }

    internal static class DefaultCSharpTranslation
    {
        public const char NewLineNormalized = '\n';
        internal static string GetTranslatedProgramCode(CultureInfo culture, string vbsSource, NonNullImmutableList<string> externalDependencies)
        {
            var stmts = Skrypton.CSharpWriter.DefaultTranslator.TranslateExecutable(culture, vbsSource, externalDependencies);

            StringBuilder tb = new StringBuilder();
            foreach (var s in stmts)
            {
                if (!s.HasContent)
                {
                    tb.Append(s.Content); // no indention for blank lines
                    tb.Append(NewLineNormalized);
                }
                else
                {
                    if (s.IndentationDepth > 0)
                    {
                        tb.Append(new string(' ', s.IndentationDepth * 4));
                    }

                    tb.Append(s.Content);
                    tb.Append(NewLineNormalized);
                }
            }
            string csText = tb.ToString();
            return csText;
        }
    }
}
