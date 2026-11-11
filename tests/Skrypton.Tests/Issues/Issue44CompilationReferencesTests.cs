using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/44 - the compilation references must not be loaded by simple name (that fails on .NET Framework)
    /// Note: The .NET Framework 4.8 case itself was verified manually with a net48 x86 console host (CI runs on Linux and has no .NET Framework).
    /// </summary>
    [TestClass]
    public sealed class Issue44CompilationReferencesTests : TestBase
    {
        // RoslynScriptControl is internal to Skrypton, so it is reached through reflection
        private static readonly Type RoslynScriptControlType = typeof(ScriptControlClass).Assembly.GetType("Skrypton.ScriptControlSupport.RoslynScriptControl", throwOnError: true)!;

        private static string[] FacadeAssemblyNames =>
            (string[])RoslynScriptControlType.GetField("FacadeAssemblyNames", BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null)!;

        private static MetadataReference[] GetCompilationReferences() =>
            (MetadataReference[])RoslynScriptControlType.GetMethod("GetCompilationReferences", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, null)!;

        [TestMethod]
        public void FacadesAreRequestedByFullStrongName()
        {
            foreach (string facadeAssemblyName in FacadeAssemblyNames)
            {
                AssemblyName name = new AssemblyName(facadeAssemblyName);
                Assert.IsNotNull(name.Version, facadeAssemblyName);
                Assert.IsNotNull(name.GetPublicKeyToken(), facadeAssemblyName);
                Assert.AreNotEqual(0, name.GetPublicKeyToken()!.Length, facadeAssemblyName);
            }
        }

        [TestMethod]
        public void ReferencesIncludeNetStandardAndTheRuntimeSupportAssembly()
        {
            string[] referencePaths = GetCompilationReferences()
                .OfType<PortableExecutableReference>()
                .Select(r => r.FilePath!)
                .ToArray();

            Assert.IsTrue(referencePaths.All(File.Exists), string.Join(Environment.NewLine, referencePaths));
            Assert.IsTrue(referencePaths.Any(p => Path.GetFileName(p).Equals("netstandard.dll", StringComparison.OrdinalIgnoreCase)), string.Join(Environment.NewLine, referencePaths));
            Assert.IsTrue(referencePaths.Any(p => Path.GetFileName(p).Equals("Skrypton.dll", StringComparison.OrdinalIgnoreCase)), string.Join(Environment.NewLine, referencePaths));
            Assert.HasCount(referencePaths.Distinct(StringComparer.OrdinalIgnoreCase).Count(), referencePaths, "no duplicates");
        }
    }
}
