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
        internal static string GetTranslatedProgramCode(CultureInfo culture, string vbsSource, IReadOnlyCollection<string> externalDependencies)
        {
            return Skrypton.CSharpWriter.DefaultTranslator.TranslateExecutable(culture, vbsSource, externalDependencies);
        }
    }
}
