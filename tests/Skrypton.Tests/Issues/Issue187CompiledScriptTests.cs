#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/187 - the whole script is parsed, translated and compiled once per version (keyed by the source,
    /// not by the translated C#) and reused for Procedures, ExecuteStatement of a procedure call and Run
    /// </summary>
    [TestClass]
    [DoNotParallelize] // the compiled scripts are cached process-wide and one of these tests changes CompiledProgramCache.MaxSize
    public sealed class Issue187CompiledScriptTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        // every test compiles its own script text, so that the tests do not depend on what other tests cached
        private static string Script(string body) => "' " + Guid.NewGuid().ToString("N") + "\r\n" + body;

        private const string Procedures =
            "Dim g\r\n" +
            "g = 10\r\n" +
            "Sub Report()\r\n" +
            "    collector.Report g + 1\r\n" +
            "End Sub\r\n" +
            "Function Add(a, b)\r\n" +
            "    Add = a + b + g\r\n" +
            "End Function\r\n" +
            "Sub Twice(ByRef x)\r\n" +
            "    x = x * 2\r\n" +
            "End Sub\r\n" +
            "Function Divide(a)\r\n" +
            "    Divide = a / 0\r\n" +
            "End Function\r\n";

        private (ScriptControlClass Class, IScriptControl Control, ValueCollector Collector) Create(string scriptContent)
        {
            ScriptControlClass scriptControlClass = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            IScriptControl sc = scriptControlClass;
            sc.AddCode(scriptContent);
            var collector = new ValueCollector();
            sc.AddObject("collector", collector, false);
            return (scriptControlClass, sc, collector);
        }

        private static object Run(IScriptControl sc, string procedureName, params object[] arguments) => sc.Run(procedureName, ref arguments);

        [TestMethod]
        public void UnchangedScriptIsCompiledOnceAndReused()
        {
            string scriptContent = Script(Procedures);
            var first = Create(scriptContent);
            var second = Create(scriptContent);

            Assert.AreEqual(0, first.Control.Procedures["Report"].NumArgs);
            first.Control.ExecuteStatement("Report");
            Assert.AreEqual(0, second.Control.Procedures["Report"].NumArgs);
            second.Control.ExecuteStatement("Report");

            Assert.IsFalse(first.Class.LastCompiledScriptReused);
            Assert.IsTrue(second.Class.LastCompiledScriptReused, "an unchanged script must not be parsed, translated or compiled again");
            Assert.AreEqual(first.Class.LastCompiledScriptKey, second.Class.LastCompiledScriptKey);
            Assert.AreSame(first.Class.LastExecutedProgramAssembly, second.Class.LastExecutedProgramAssembly);
            Assert.AreSame(first.Control.Procedures, second.Control.Procedures, "the procedures of an unchanged script are parsed once");
            Assert.AreEqual(11, Convert.ToInt32(first.Collector.Values.Single()));
            Assert.AreEqual(11, Convert.ToInt32(second.Collector.Values.Single()), "each run has its own state");
        }

        [TestMethod]
        public void ChangedScriptIsCompiledAgainAndTheNewCodeRuns()
        {
            string version1 = Script(Procedures);
            var first = Create(version1);
            first.Control.ExecuteStatement("Call Report()");
            var changed = Create(version1.Replace("g + 1", "g + 2"));
            changed.Control.ExecuteStatement("Call Report()");

            Assert.IsFalse(changed.Class.LastCompiledScriptReused);
            Assert.AreNotEqual(first.Class.LastCompiledScriptKey, changed.Class.LastCompiledScriptKey);
            Assert.AreEqual(12, Convert.ToInt32(changed.Collector.Values.Single()));
        }

        [TestMethod]
        public void RunCallsProceduresWithArguments()
        {
            var created = Create(Script(Procedures));

            Assert.AreEqual(15, Convert.ToInt32(Run(created.Control, "Add", 2, 3)));
            object[] arguments = [21];
            created.Control.Run("Twice", ref arguments);
            Assert.IsTrue(created.Class.LastCompiledScriptReused, "Run reuses the script compiled by the previous call");
            Assert.AreEqual(42, Convert.ToInt32(arguments[0]), "a ByRef argument is updated");
            Assert.ThrowsExactly<ArgumentException>(() => Run(created.Control, "Missing"));
        }

        [TestMethod]
        public void RuntimeErrorIsMappedToTheVBScriptLine()
        {
            var created = Create(Script(Procedures));

            Assert.Throws<Exception>(() => Run(created.Control, "Divide", 1));
            Assert.AreEqual(11, created.Control.Error!.Number, "Division by zero");
            Assert.AreEqual(14, created.Control.Error.Line, "the line of 'Divide = a / 0' (line 1 is the comment of the test)");
            Assert.AreEqual("Divide = a / 0", created.Control.Error.Text);

            var other = Create(Script("Sub Fail()\r\n    Dim x\r\n    x = 1 / 0\r\nEnd Sub\r\n"));
            Assert.Throws<Exception>(() => other.Control.ExecuteStatement("Fail"));
            Assert.IsNotNull(other.Class.LastCompiledScriptKey);
            Assert.AreEqual(4, other.Control.Error!.Line);
        }

        [TestMethod]
        public void ParallelRunsOnTwoVersionsDoNotAffectEachOther()
        {
            string version1 = Script(Procedures);
            string[] versions = [version1, version1.Replace("g = 10", "g = 20")];
            int[] results = Task.WhenAll(Enumerable.Range(0, 40).Select(i => Task.Run(() => Convert.ToInt32(Run(Create(versions[i % 2]).Control, "Add", i, 0))))).GetAwaiter().GetResult();

            for (int i = 0; i < results.Length; i++)
                Assert.AreEqual(i + ((i % 2 == 0) ? 10 : 20), results[i]);
        }

        [TestMethod]
        public void ParallelCallersOfANewVersionCompileItOnce()
        {
            string scriptContent = Script(Procedures);
            bool[] reused = Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() =>
            {
                var created = Create(scriptContent);
                created.Control.ExecuteStatement("Report");
                return created.Class.LastCompiledScriptReused;
            }))).GetAwaiter().GetResult();

            Assert.AreEqual(1, reused.Count(r => !r));
        }

        [TestMethod]
        public void StatementsOtherThanAProcedureCallAreTranslatedWithTheScript()
        {
            var created = Create(Script(Procedures));
            created.Control.ExecuteStatement("collector.Report Add(1, 2)");
            Assert.IsNull(created.Class.LastCompiledScriptKey);
            Assert.AreEqual(13, Convert.ToInt32(created.Collector.Values.Single()));
        }

        [TestMethod]
        public void CachingCanBeDisabled()
        {
            int originalMaxSize = CompiledProgramCache.MaxSize;
            try
            {
                CompiledProgramCache.MaxSize = 0;
                string scriptContent = Script(Procedures);
                Create(scriptContent).Control.ExecuteStatement("Report");
                var second = Create(scriptContent);
                second.Control.ExecuteStatement("Report");
                Assert.IsFalse(second.Class.LastCompiledScriptReused, "with caching disabled every call compiles");
            }
            finally
            {
                CompiledProgramCache.MaxSize = originalMaxSize;
            }
        }
    }
}
