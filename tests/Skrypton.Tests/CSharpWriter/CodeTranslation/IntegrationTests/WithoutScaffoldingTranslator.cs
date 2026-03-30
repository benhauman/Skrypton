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
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    internal static class WithoutScaffoldingTranslator
    {
        public static NonNullImmutableList<string> DefaultConsoleExternalDependencies = new NonNullImmutableList<string>().Add("WScript");
    }
    internal static class DefaultCSharpTranslation
    {
        internal static string GetTranslatedProgramCode(TestBaseX tst, string vbsSource, IReadOnlyCollection<string> externalDependencies, IReadOnlyCollection<ExternalMemberMethodInfo> externalMemberMethods, string[] translationSuppression)
        {
            var scriptengineClass = tst.CreateScriptControlClass(new TestRuntimeHost(tst.CreateTestHostServices()), translationSuppression);
            IScriptControl scriptengine = scriptengineClass;
            foreach (string externalDependencyName in externalDependencies)
            {
                var addMembers = externalMemberMethods.Any(m => m.OwnerName == externalDependencyName);
                if (addMembers)
                {
                    scriptengine.AddObject(externalDependencyName, new object(), AddMembers: false); // added below
                }
                else
                {
                    scriptengine.AddObject(externalDependencyName, new object(), AddMembers: false);
                }
            }

            Dictionary<string, string[]> dictMembers =
                externalMemberMethods
                    .GroupBy(x => x.OwnerName)
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(x => x.MethodName).ToArray());
            return scriptengineClass.TestGenerateCSharpCode(vbsSource, dictMembers);
        }
    }
}
