#nullable enable
using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in WScript.Shell on the host's IHostProcessControlHostService / IHostEnvironmentHostService
    /// </summary>
    [TestClass]
    public sealed class ScriptonWScriptShellTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private sealed class TestEnvironment : IHostEnvironmentHostService
        {
            public string ExpandEnvironmentStrings(string source) => source.Replace("%TEMP%", "/tmp-test", StringComparison.OrdinalIgnoreCase);
        }

        private List<object> Execute(string vbs, TestHostProcessControlHostService? processControl, IHostEnvironmentHostService? environment = null)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (processControl != null)
                        s.RegisterHostService<IHostProcessControlHostService>(() => processControl);
                    if (environment != null)
                        s.RegisterHostService<IHostEnvironmentHostService>(() => environment);
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("Dim sh : Set sh = CreateObject(\"WScript.Shell\")\r\n" + vbs);
            return collector.Values;
        }

        [TestMethod]
        public void RunGoesThroughTheHostService()
        {
            var processControl = new TestHostProcessControlHostService();
            var values = Execute(
                "collector.Report sh.Run(\"notepad.exe\")\r\n" +
                "sh.Run \"test.txt\", 0, True\r\n" +
                "collector.Report TypeName(sh)\r\n",
                processControl);
            CollectionAssert.AreEqual(new object[] { 0, "IWshShell3" }, values.ToArray());
            CollectionAssert.AreEqual(new[] { "notepad.exe|1|False", "test.txt|0|True" }, processControl.Started.ToArray());
        }

        [TestMethod]
        public void ExecReturnsTheExitCodeAndOutput()
        {
            var values = Execute(
                "Dim ex : Set ex = sh.Exec(\"cmd /c dir\")\r\n" +
                "collector.Report ex.Status & \"|\" & ex.ExitCode & \"|\" & ex.ProcessID\r\n" +
                "collector.Report ex.StdOut.ReadLine()\r\n" +
                "collector.Report ex.StdOut.ReadAll()\r\n" +
                "collector.Report ex.StdOut.AtEndOfStream\r\n",
                new TestHostProcessControlHostService());
            CollectionAssert.AreEqual(new object[] { "1|0|4711", "out line 1", "out line 2\r\n", true }, values.ToArray());
        }

        [TestMethod]
        public void WithoutAHostServiceRunRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute("sh.Run \"notepad.exe\"\r\n", processControl: null));
            StringAssert.Contains(exception.ToString(), nameof(IHostProcessControlHostService));
        }

        [TestMethod]
        public void ExpandEnvironmentStrings()
        {
            Assert.AreEqual("/tmp-test\\x", Execute("collector.Report sh.ExpandEnvironmentStrings(\"%TEMP%\\x\")\r\n", null, new TestEnvironment())[0]);
            Assert.AreEqual(Environment.ExpandEnvironmentVariables("%PATH%"), Execute("collector.Report sh.ExpandEnvironmentStrings(\"%PATH%\")\r\n", null)[0]);
        }
    }
}
