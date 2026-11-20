#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// Issue #190: the compiled versions of IEagerScriptEngine are loaded through the code load context of WindowsRuntimeHost - on .NET a collectible
    /// AssemblyLoadContext that is unloaded once the version is released, on the .NET Framework the default context (release only drops the reference)
    /// </summary>
    [TestClass]
    public sealed class EagerScriptEngineCodeLoadContextTests
    {
        public sealed class ObjectCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        private sealed class Logger : IRuntimeLogger
        {
            public void LogException(Exception exception)
            {
            }
        }

        // the script reports an instance of its own class - its type lives in the assembly of the compiled version
        private static string Script() =>
            "' " + Guid.NewGuid().ToString("N") + "\r\n" +
            "Class Thing\r\n" +
            "End Class\r\n" +
            "Function Answer()\r\n" +
            "    collector.Report New Thing\r\n" +
            "    Answer = 42\r\n" +
            "End Function\r\n";

        private static readonly ScriptEngineConfiguration Configuration = new ScriptEngineConfiguration(false, null, false, Array.Empty<string>(), Array.Empty<string>(), mode: ScriptEngineMode.EagerScript);

        [TestMethod]
        public void TheDefaultHostHasACodeLoadContext()
        {
            Assert.IsInstanceOfType<WindowsHostCodeLoadContextHostService>(WindowsRuntimeHost.CreateDefault().TryGetRuntimeHostService<IHostCodeLoadContextHostService>());
        }

        [TestMethod]
        public void AReleasedVersionIsUnloaded()
        {
            WindowsRuntimeHost host = WindowsRuntimeHost.CreateDefault();
            CompiledCodeHandle handle = Compile(host, Script());
            IEagerScriptEngine engine = handle.CreateEngine(host, new Logger());
            WeakReference programAssembly = RunAndGetProgramAssembly(engine);

            handle.Dispose(); // retired, but the engine still runs
            Assert.IsTrue(programAssembly.IsAlive);
            engine.Dispose(); // the last run of the version is done

#if NETFRAMEWORK
            Assert.IsTrue(programAssembly.IsAlive, "the .NET Framework can not unload an assembly: release only drops the reference");
#else
            for (int i = 0; (i < 20) && programAssembly.IsAlive; i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            Assert.IsFalse(programAssembly.IsAlive, "the collectible code load context of the released version is unloaded");
#endif
        }

        [MethodImpl(MethodImplOptions.NoInlining)] // so that nothing of the run stays on the stack of the test
        private static CompiledCodeHandle Compile(WindowsRuntimeHost host, string code)
        {
            IEagerScriptEngine engine = new EagerScriptEngine(host, new Logger(), CultureInfo.InvariantCulture, Configuration);
            engine.DeclareObject("collector", typeof(ObjectCollector));
            engine.SetCode(code);
            CompiledCodeHandle handle = engine.CompileCode();
            engine.Dispose();
            return handle;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static WeakReference RunAndGetProgramAssembly(IEagerScriptEngine engine)
        {
            ObjectCollector collector = new ObjectCollector();
            engine.BindObject("collector", collector);
            Assert.AreEqual(42, Convert.ToInt32(engine.CallProcedure("Answer"), CultureInfo.InvariantCulture));
            return new WeakReference(collector.Values[0]!.GetType().Assembly);
        }
    }
}
