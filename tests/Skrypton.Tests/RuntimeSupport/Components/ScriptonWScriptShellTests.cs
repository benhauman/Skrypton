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

        /// <summary>
        /// An environment kept by a host's IHostEnvironmentVariablesHostService
        /// </summary>
        private sealed class TestEnvironmentVariables : IHostEnvironmentVariablesHostService
        {
            public Dictionary<string, Dictionary<string, string>> Types { get; } = new Dictionary<string, Dictionary<string, string>>
            {
                ["PROCESS"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["ProgramFiles"] = @"C:\Program Files", ["OS"] = "Windows_NT" },
                ["SYSTEM"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["OS"] = "Windows_NT", ["Path"] = @"%SystemRoot%\system32" },
                ["USER"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["TEMP"] = @"%USERPROFILE%\AppData\Local\Temp" },
                ["VOLATILE"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            };

            public IReadOnlyList<KeyValuePair<string, string>> GetVariables(string type) => new List<KeyValuePair<string, string>>(Types[type]);

            public string? GetVariable(string type, string name) => Types[type].TryGetValue(name, out string? value) ? value : null;

            public void SetVariable(string type, string name, string? value)
            {
                if (value == null)
                    Types[type].Remove(name);
                else
                    Types[type][name] = value;
            }
        }

        private List<object> ExecuteWithVariables(string vbs, IHostEnvironmentVariablesHostService variables)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostEnvironmentVariablesHostService>(() => variables))),
                CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("Dim sh : Set sh = CreateObject(\"WScript.Shell\")\r\n" + vbs);
            return collector.Values;
        }

        /// <summary>
        /// WshShell.Environment(type) through the host's IHostEnvironmentVariablesHostService (issue #107): env(name) reads ("" if there is none),
        /// env(name) = value sets, Remove removes (no error if there is none), Count / length, For Each over "NAME=value", SYSTEM by default,
        /// the type case-insensitive; an unknown type and an empty name are error 5
        /// </summary>
        [TestMethod]
        public void EnvironmentGoesThroughTheHostService()
        {
            var variables = new TestEnvironmentVariables();
            var values = ExecuteWithVariables(
                "Dim e : Set e = sh.Environment(\"PROCESS\")\r\n" +
                "collector.Report TypeName(e) & \"|\" & e(\"ProgramFiles\") & \"|[\" & e(\"NO_SUCH_VARIABLE\") & \"]|\" & e.Item(\"os\")\r\n" +
                "e(\"SKRYPTON_PROBE\") = \"1\"\r\n" +
                "collector.Report e.Count & \"|\" & e.length & \"|\" & sh.Environment(\"process\")(\"SKRYPTON_PROBE\")\r\n" +
                "Dim v, s\r\n" +
                "For Each v In e\r\n" +
                "    s = s & v & \";\"\r\n" +
                "Next\r\n" +
                "collector.Report s\r\n" +
                "e.Remove \"SKRYPTON_PROBE\"\r\n" +
                "e.Remove \"SKRYPTON_NOT_THERE\"\r\n" +
                "collector.Report e.Count\r\n" +
                "collector.Report sh.Environment()(\"Path\") & \"|\" & sh.Environment(\"USER\")(\"TEMP\") & \"|\" & sh.Environment(\"VOLATILE\").Count\r\n" +
                "Dim x : Set x = sh.Environment\r\n" +
                "collector.Report x(\"OS\")\r\n" +
                "On Error Resume Next\r\n" +
                "Set x = sh.Environment(\"BOGUS\")\r\n" +
                "collector.Report Err.Number\r\n" +
                "Err.Clear\r\n" +
                "e(\"\") = \"x\"\r\n" +
                "collector.Report Err.Number\r\n",
                variables);
            CollectionAssert.AreEqual(
                new object[]
                {
                    @"IWshEnvironment|C:\Program Files|[]|Windows_NT",
                    "3|3|1",
                    @"ProgramFiles=C:\Program Files;OS=Windows_NT;SKRYPTON_PROBE=1;",
                    2,
                    @"%SystemRoot%\system32|%USERPROFILE%\AppData\Local\Temp|0",
                    "Windows_NT",
                    5,
                    5,
                },
                values.ToArray());
            Assert.IsFalse(variables.Types["PROCESS"].ContainsKey("SKRYPTON_PROBE"));
        }

        /// <summary>
        /// A host that sandboxes ExpandEnvironmentStrings with its own IHostEnvironmentHostService (and no IHostEnvironmentVariablesHostService):
        /// Environment("PROCESS") reads through it and does not change the process environment (issue #107)
        /// </summary>
        [TestMethod]
        public void EnvironmentOfASandboxedHostReadsThroughItsExpandEnvironmentStrings()
        {
            var values = Execute(
                "Dim e : Set e = sh.Environment(\"PROCESS\")\r\n" +
                "collector.Report e(\"TEMP\") & \"|[\" & e(\"PATH\") & \"]\"\r\n" +
                "On Error Resume Next\r\n" +
                "e(\"SKRYPTON_SANDBOX\") = \"1\"\r\n" +
                "collector.Report (Err.Number <> 0)\r\n",
                null,
                new TestEnvironment());
            CollectionAssert.AreEqual(new object[] { "/tmp-test|[]", true }, values.ToArray());
            Assert.IsNull(Environment.GetEnvironmentVariable("SKRYPTON_SANDBOX"));
        }

        /// <summary>
        /// Without host services: the process environment
        /// </summary>
        [TestMethod]
        public void EnvironmentOfTheProcess()
        {
            Environment.SetEnvironmentVariable("SKRYPTON_ENV_TEST", "before");
            try
            {
                var values = Execute(
                    "Dim e : Set e = sh.Environment(\"PROCESS\")\r\n" +
                    "collector.Report e(\"SKRYPTON_ENV_TEST\")\r\n" +
                    "e(\"SKRYPTON_ENV_TEST\") = \"after\"\r\n" +
                    "collector.Report sh.ExpandEnvironmentStrings(\"%SKRYPTON_ENV_TEST%\")\r\n" +
                    "collector.Report (e.Count > 0)\r\n",
                    null);
                CollectionAssert.AreEqual(new object[] { "before", "after", true }, values.ToArray());
            }
            finally
            {
                Environment.SetEnvironmentVariable("SKRYPTON_ENV_TEST", null);
            }
        }
    }
}
