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
    /// https://github.com/benhauman/Skrypton/issues/190 - IEagerScriptEngine: declare objects, set and compile the code once (no instances, nothing
    /// runs), a disposable compiled code handle, then bind and call procedures / run the global code once per engine
    /// </summary>
    [TestClass]
    public sealed class Issue190EagerScriptEngineTests : TestBase
    {
        public interface IReporter
        {
            void Report(object? value);
        }

        public sealed class ValueCollector : IReporter
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        // every test compiles its own script text, so that the tests do not depend on versions that other tests compiled
        private static string Script(string body) => "' " + Guid.NewGuid().ToString("N") + "\r\n" + body;

        private const string Counter =
            "Dim n\r\n" +
            "n = 100\r\n" +
            "collector.Report \"global\"\r\n" +
            "Function Inc()\r\n" +
            "    n = n + 1\r\n" +
            "    Inc = n\r\n" +
            "End Function\r\n" +
            "Sub Twice(ByRef x)\r\n" +
            "    x = x * 2\r\n" +
            "End Sub\r\n" +
            "Function Divide(a)\r\n" +
            "    Divide = a / 0\r\n" +
            "End Function\r\n";

        private static ScriptEngineConfiguration EagerConfiguration() => new ScriptEngineConfiguration(false, null, false, [], [], mode: ScriptEngineMode.EagerScript);

        private IEagerScriptEngine CreateEngine() => new EagerScriptEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger, TestCulture, EagerConfiguration());

        private CompiledCodeHandle Compile(string code, out IEagerScriptEngine engine, Type? collectorType = null)
        {
            engine = CreateEngine();
            engine.DeclareObject("collector", collectorType ?? typeof(ValueCollector));
            engine.SetCode(code);
            return engine.CompileCode();
        }

        [TestMethod]
        public void CompileNeedsNoInstancesAndRunsNothing()
        {
            using CompiledCodeHandle handle = Compile(Script(Counter), out IEagerScriptEngine engine);
            using (engine)
            {
                Assert.IsFalse(handle.Reused);
                Assert.AreEqual(64, handle.Key.Length, "SHA-256");
                CollectionAssert.AreEquivalent(new[] { "Inc", "Twice", "Divide" }, handle.Procedures.Select(p => p.Name).ToArray());
                Assert.IsTrue(engine.ContainsProcedure("inc"));

                var collector = new ValueCollector();
                engine.BindObject("collector", collector);
                Assert.AreEqual(0, collector.Values.Count, "binding runs nothing");
                Assert.AreEqual(101, Convert.ToInt32(engine.CallProcedure("Inc")));
                Assert.AreEqual(102, Convert.ToInt32(engine.CallProcedure("Inc")), "the module state is kept between calls");
                CollectionAssert.AreEqual(new object?[] { "global" }, collector.Values, "the global code runs once per engine (and not at compile)");
            }
        }

        [TestMethod]
        public void SecondCompileOfTheSameVersionIsReused()
        {
            string code = Script(Counter);
            using CompiledCodeHandle first = Compile(code, out IEagerScriptEngine firstEngine);
            using CompiledCodeHandle second = Compile(code, out IEagerScriptEngine secondEngine);
            firstEngine.Dispose();
            secondEngine.Dispose();

            Assert.IsFalse(first.Reused);
            Assert.IsTrue(second.Reused, "no parse, no translation, no compile");
            Assert.AreEqual(first.Key, second.Key);
            Assert.AreNotEqual(first.Key, Compile(code, out IEagerScriptEngine other, typeof(IReporter)).Key, "the declared type is part of the key");
            other.Dispose();
        }

        [TestMethod]
        public void ParallelFirstCompileOfAVersionCompilesOnce()
        {
            string code = Script(Counter);
            CompiledCodeHandle[] handles = Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() =>
            {
                CompiledCodeHandle handle = Compile(code, out IEagerScriptEngine engine);
                engine.Dispose();
                return handle;
            }))).GetAwaiter().GetResult();
            try
            {
                Assert.AreEqual(1, handles.Count(h => !h.Reused));
                Assert.AreEqual(1, handles.Select(h => h.Key).Distinct().Count());
            }
            finally
            {
                foreach (CompiledCodeHandle handle in handles)
                    handle.Dispose();
            }
        }

        [TestMethod]
        public void EnginesFromAHandleRunTheGlobalCodeOnceEach()
        {
            using CompiledCodeHandle handle = Compile(Script(Counter), out IEagerScriptEngine compiling);
            compiling.Dispose();

            for (int run = 0; run < 2; run++)
            {
                using IEagerScriptEngine engine = handle.CreateEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger);
                var collector = new ValueCollector();
                engine.BindObject("collector", collector);
                engine.RunGlobalCode();
                Assert.AreEqual(101, Convert.ToInt32(engine.CallProcedure("Inc")), "every engine has its own module state");
                Assert.AreEqual(102, Convert.ToInt32(engine.CallProcedure("Inc")));
                object?[] args = [21];
                engine.CallProcedure("Twice", args);
                Assert.AreEqual(42, Convert.ToInt32(args[0]), "a ByRef argument is updated");
                Assert.AreEqual(1, collector.Values.Count);
                Assert.ThrowsExactly<InvalidOperationException>(() => engine.RunGlobalCode(), "the global code runs once");
            }
        }

        [TestMethod]
        public void TopLevelScriptRunsWithRunGlobalCode()
        {
            string code = Script(
                "Dim total\r\n" +
                "total = Twice(4)\r\n" +
                "collector.Report total\r\n" +
                "Function Twice(x)\r\n" +
                "    Twice = x * 2\r\n" +
                "End Function\r\n" +
                "Function GetTotal()\r\n" +
                "    GetTotal = total\r\n" +
                "End Function\r\n");
            using CompiledCodeHandle handle = Compile(code, out IEagerScriptEngine compiling);
            compiling.Dispose();
            using CompiledCodeHandle again = Compile(code, out IEagerScriptEngine second);
            second.Dispose();
            Assert.IsTrue(again.Reused);

            using IEagerScriptEngine engine = again.CreateEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger);
            var collector = new ValueCollector();
            engine.BindObject("collector", collector);
            engine.RunGlobalCode();
            Assert.AreEqual(8, Convert.ToInt32(collector.Values.Single()));
            Assert.AreEqual(8, Convert.ToInt32(engine.CallProcedure("GetTotal")), "the global variables are kept for the procedures");
        }

        [TestMethod]
        public void CallsOutOfOrderAreRefused()
        {
            using IEagerScriptEngine engine = CreateEngine();
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.CompileCode());
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.BindObject("collector", new ValueCollector()));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.CallProcedure("Inc"));
            Assert.ThrowsExactly<InvalidOperationException>(() => _ = engine.Procedures);
            engine.DeclareObject("collector", typeof(IReporter));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.DeclareObject("Collector", typeof(object)), "each name once");
            engine.SetCode(Script(Counter));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.SetCode("x = 1"));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.DeclareObject("other", typeof(object)));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.RunGlobalCode());

            using CompiledCodeHandle handle = engine.CompileCode();
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.CompileCode(), "sealed");
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.SetCode("x = 1"), "sealed");
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.RunGlobalCode(), "every declared object must be bound");
            Assert.ThrowsExactly<ArgumentException>(() => engine.BindObject("other", new ValueCollector()), "not declared");
            Assert.ThrowsExactly<ArgumentException>(() => engine.BindObject("collector", "not a reporter"), "wrong type");
            Assert.ThrowsExactly<ArgumentNullException>(() => engine.BindObject("collector", null!));
            engine.BindObject("collector", new ValueCollector()); // an interface passes for an implementing class
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.BindObject("collector", new ValueCollector()), "bound once");
            Assert.ThrowsExactly<ArgumentException>(() => engine.CallProcedure("Missing"));
            Assert.AreEqual(101, Convert.ToInt32(engine.CallProcedure("Inc")));
            Assert.ThrowsExactly<InvalidOperationException>(() => engine.BindObject("collector", new ValueCollector()), "the run has started");

            engine.Dispose();
            Assert.ThrowsExactly<ObjectDisposedException>(() => engine.CallProcedure("Inc"));
            handle.Dispose();
            Assert.ThrowsExactly<ObjectDisposedException>(() => handle.CreateEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger));
        }

        [TestMethod]
        public void RuntimeErrorOfAnEngineFromAHandleIsMappedToTheVBScriptLine()
        {
            using CompiledCodeHandle handle = Compile(Script(Counter), out IEagerScriptEngine compiling);
            compiling.Dispose();
            using IEagerScriptEngine engine = handle.CreateEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger);
            engine.BindObject("collector", new ValueCollector());

            Assert.Throws<Exception>(() => engine.CallProcedure("Divide", 1));
            Assert.AreEqual(11, engine.Error.Number, "Division by zero");
            Assert.AreEqual(13, engine.Error.Line, "the line of 'Divide = a / 0' (line 1 is the comment of the test)");
            Assert.AreEqual("Divide = a / 0", engine.Error.Text);
            Assert.AreEqual(101, Convert.ToInt32(engine.CallProcedure("Inc")), "the engine still runs after a runtime error of a procedure");
            Assert.AreEqual(0, engine.Error.Number, "a call that succeeds clears the error");
        }

        [TestMethod]
        public void CompileErrorIsRaisedByCompileCodeAndNotCached()
        {
            string code = Script("Sub Broken(\r\nEnd Sub\r\n");
            using IEagerScriptEngine engine = CreateEngine();
            engine.SetCode(code);
            Assert.Throws<Exception>(() => engine.CompileCode());
            Assert.AreNotEqual(0, engine.Error.Number);
            Assert.Throws<Exception>(() => engine.CompileCode(), "a failed compilation is not cached, so it compiles (and fails) again");
        }

        [TestMethod]
        public void AReleasedVersionIsCompiledAgain()
        {
            string code = Script(Counter);
            CompiledCodeHandle handle = Compile(code, out IEagerScriptEngine engine);
            IEagerScriptEngine fromHandle = handle.CreateEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger);
            handle.Dispose();
            engine.Dispose();
            using (CompiledCodeHandle stillUsed = Compile(code, out IEagerScriptEngine second))
            {
                second.Dispose();
                Assert.IsTrue(stillUsed.Reused, "a retired version is kept while an engine of it runs");
            }
            fromHandle.Dispose();

            using CompiledCodeHandle compiledAgain = Compile(code, out IEagerScriptEngine third);
            third.Dispose();
            Assert.IsFalse(compiledAgain.Reused, "the version was released with its last handle and engine");
        }

        [TestMethod]
        public void TheModeSelectsTheEngine()
        {
            ScriptEngineConfiguration scriptControl = new ScriptEngineConfiguration(false, null, false, [], []);
            Assert.AreEqual(ScriptEngineMode.ScriptControl, scriptControl.Mode);
            Assert.ThrowsExactly<ArgumentException>(() => new EagerScriptEngine(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger, TestCulture, scriptControl));
            Assert.ThrowsExactly<ArgumentException>(() => CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), EagerConfiguration()));
            Assert.IsNotNull(CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), scriptControl), "ScriptControlClass is unchanged");
        }
    }
}
