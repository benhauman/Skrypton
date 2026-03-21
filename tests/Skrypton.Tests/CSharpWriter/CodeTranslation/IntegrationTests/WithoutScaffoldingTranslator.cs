using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using System;
using System.Globalization;
using System.Linq;
using Skrypton.CSharpWriter.CodeTranslation;
using System.Text;
using Microsoft.Testing.Extensions.VSTestBridge;
using System.Collections.Generic;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    internal static class WithoutScaffoldingTranslator
    {
        public static NonNullImmutableList<string> DefaultConsoleExternalDependencies = new NonNullImmutableList<string>().Add("WScript");
    }
    internal static class DefaultCSharpTranslation
    {
        public const char NewLineNormalized = '\n';
        internal static string GetTranslatedProgramCode(CultureInfo culture, string vbsSource, IReadOnlyCollection<string> externalDependencies)
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
                    s.RenderTranslatedStatement(tb);
                    tb.Append(NewLineNormalized);
                }
            }
            string csText = tb.ToString();
            return csText;
        }
    }
}
