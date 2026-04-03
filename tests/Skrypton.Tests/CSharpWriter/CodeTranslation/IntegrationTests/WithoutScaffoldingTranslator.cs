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
using Skrypton.Tests.Application;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    internal static class WithoutScaffoldingTranslator
    {
        public static NonNullImmutableList<string> DefaultConsoleExternalDependencies = new NonNullImmutableList<string>().Add("WScript");
    }
    internal static class DefaultCSharpTranslation
    {
        internal static string GetTranslatedProgramCode(TestBaseX tst, string vbsSource, IReadOnlyCollection<string> externalDependencies, IReadOnlyCollection<ExternalMemberMethodInfo> externalMemberMethods, string[] translationSuppression, string[] noWarn)
        {
            Dictionary<string, ScriptExternalReferenceInfo> xr = new Dictionary<string, ScriptExternalReferenceInfo>();
            foreach (string externalDependencyName in externalDependencies)
            {
                string[] members = externalMemberMethods.Where(m => m.OwnerName == externalDependencyName).Select(x => x.MethodName).ToArray();
                xr.Add(externalDependencyName, new ScriptExternalReferenceInfo(instance: new object(), members));
            }
            //Dictionary<string, string[]> dictMembers =
            //    externalMemberMethods
            //        .GroupBy(x => x.OwnerName)
            //        .ToDictionary(
            //            g => g.Key,
            //            g => g.Select(x => x.MethodName).ToArray());

            return GetTranslatedProgramCodeX(tst, vbsSource, xr, translationSuppression, noWarn);
        }
        internal static string GetTranslatedProgramCodeX(TestBaseX tst, string vbsSource,
            IReadOnlyDictionary<string, ScriptExternalReferenceInfo> externalDependencies,
            string[] translationSuppression,
            string[] noWarn)
        {
            var scriptengineConfig = tst.CreateScriptControlConfiguration(false, translationSuppression, noWarn);
            var scriptengineClass = tst.CreateScriptControlClass(new TestRuntimeHost(tst.CreateTestHostServices()), scriptengineConfig);
            IScriptControl scriptengine = scriptengineClass;
            foreach (var externalDependencyInfo in externalDependencies)
            {
                scriptengine.AddObject(externalDependencyInfo.Key, externalDependencyInfo.Value.Instance, AddMembers: externalDependencyInfo.Value.AddMembers); // added explicitly
            }
            scriptengine.AddCode(vbsSource);
            return scriptengineClass.TestGenerateCSharpCode(null, null);
        }
    }
}
