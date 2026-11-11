using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/43 - identical programs must be compiled and loaded once, not once per ExecuteStatement call
    /// </summary>
    [TestClass]
    [DoNotParallelize] // CompiledProgramCache is process-wide and some of these tests change its MaxSize
    public sealed class Issue43CompiledProgramCacheTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private (Assembly Assembly, List<object> Values) Execute(string vbs, int value)
        {
            var collector = new ValueCollector();
            ScriptControlClass scriptControlClass = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            IScriptControl scriptengine = scriptControlClass;
            scriptengine.AddObject("collector", collector, false);
            scriptengine.AddObject("input", new Input(value), false);
            scriptengine.ExecuteStatement(vbs);
            return (scriptControlClass.LastExecutedProgramAssembly!, collector.Values);
        }

        public sealed class Input
        {
            public Input(int value) => Value = value;
            public int Value { get; }
        }

        private const string Script = "Dim g : g = g + 1\r\ncollector.Report input.Value * 10 + g\r\n";

        [TestMethod]
        public void IdenticalProgramsAreCompiledOnceAndStateIsNotShared()
        {
            var first = Execute(Script, 1);
            var second = Execute(Script, 2);

            Assert.AreSame(first.Assembly, second.Assembly, "the second run must reuse the compiled program");
            // Each run gets new external references and new global state (g starts as Empty in every run)
            Assert.AreEqual(11, Convert.ToInt32(first.Values[0]));
            Assert.AreEqual(21, Convert.ToInt32(second.Values[0]));
        }

        [TestMethod]
        public void AssemblyCountStaysConstantForRepeatedRuns()
        {
            Execute(Script, 1);
            Execute(Script, 1);
            int assemblyCountBefore = AppDomain.CurrentDomain.GetAssemblies().Length;
            for (int i = 0; i < 20; i++)
                Execute(Script, i);
            Assert.AreEqual(assemblyCountBefore, AppDomain.CurrentDomain.GetAssemblies().Length);
        }

        [TestMethod]
        public void DifferentProgramsAreCompiledSeparately()
        {
            var first = Execute(Script, 1);
            var other = Execute(Script + "collector.Report 0\r\n", 1);
            Assert.AreNotSame(first.Assembly, other.Assembly);
            Assert.HasCount(2, other.Values);
        }

        [TestMethod]
        public void CacheIsBoundedAndCanBeDisabled()
        {
            int originalMaxSize = CompiledProgramCache.MaxSize;
            try
            {
                CompiledProgramCache.MaxSize = 2;
                Execute(Script + "' a\r\n", 1);
                Execute(Script + "' b\r\n", 1);
                Execute(Script + "' c\r\n", 1);
                Assert.AreEqual(2, CompiledProgramCache.Count);

                CompiledProgramCache.MaxSize = 0;
                Assert.AreEqual(0, CompiledProgramCache.Count);
                var first = Execute(Script, 1);
                var second = Execute(Script, 1);
                Assert.AreNotSame(first.Assembly, second.Assembly, "with caching disabled every run compiles the program again");

                Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => CompiledProgramCache.MaxSize = -1);
            }
            finally
            {
                CompiledProgramCache.MaxSize = originalMaxSize;
            }
        }
    }
}
