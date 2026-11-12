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

        private sealed class TestEventLog : IHostEventLogHostService
        {
            public List<string> Entries { get; } = new List<string>();

            public void LogEvent(int type, string message, string? target) => Entries.Add(type + "|" + message + "|" + (target ?? "(local)"));
        }

        /// <summary>
        /// WshShell.LogEvent (issue #116) through the host's IHostEventLogHostService: the types 0, 1, 2, 4, 8, 16 (others are error 5), True as
        /// the result, the optional target computer; without a host service it returns True too (System.Diagnostics.Trace)
        /// </summary>
        [TestMethod]
        public void LogEventGoesThroughTheHostService()
        {
            var eventLog = new TestEventLog();
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostEventLogHostService>(() => eventLog))),
                CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(
                "Dim sh : Set sh = CreateObject(\"WScript.Shell\")\r\n" +
                "sh.LogEvent 2, \"warning\"\r\n" +
                "collector.Report sh.LogEvent(4, \"information\")\r\n" +
                "Dim t\r\n" +
                "For Each t In Array(0, 1, 8, 16)\r\n" +
                "    sh.LogEvent t, \"type\"\r\n" +
                "Next\r\n" +
                "sh.LogEvent 4, \"remote\", \"SERVER1\"\r\n" +
                "sh.LogEvent 4, \"local\", \"\"\r\n" +
                "On Error Resume Next\r\n" +
                "sh.LogEvent 3, \"bad\"\r\n" +
                "collector.Report Err.Number\r\n");
            CollectionAssert.AreEqual(new object[] { true, 5 }, collector.Values.ToArray());
            CollectionAssert.AreEqual(
                new[] { "2|warning|(local)", "4|information|(local)", "0|type|(local)", "1|type|(local)", "8|type|(local)", "16|type|(local)", "4|remote|SERVER1", "4|local|(local)" },
                eventLog.Entries.ToArray());

            Assert.AreEqual(true, Execute("collector.Report sh.LogEvent(4, \"to the trace\")\r\n", null)[0]);
        }

        /// <summary>
        /// A registry in memory: HKEY_LOCAL_MACHINE\Software\Test with a value of each kind and no default value
        /// </summary>
        private sealed class TestRegistry : IHostRegistryHostService
        {
            public List<string> Reads { get; } = new List<string>();

            public HostRegistryValue? ReadValue(string root, string subKey, string valueName, out bool keyExists)
            {
                Reads.Add(root + "|" + subKey + "|" + valueName);
                keyExists = (root == "HKEY_LOCAL_MACHINE") && (subKey.Equals(@"Software\Test", StringComparison.OrdinalIgnoreCase) || (subKey.Length == 0));
                if (!keyExists)
                    return null;
                switch (valueName.ToUpperInvariant())
                {
                    case "SZ": return new HostRegistryValue(HostRegistryValueKind.Sz, "text");
                    case "EXPAND": return new HostRegistryValue(HostRegistryValueKind.ExpandSz, @"%USERPROFILE%\x");
                    case "DWORD": return new HostRegistryValue(HostRegistryValueKind.DWord, -60);
                    case "QWORD": return new HostRegistryValue(HostRegistryValueKind.QWord, 5L);
                    case "MULTI": return new HostRegistryValue(HostRegistryValueKind.MultiSz, new[] { "a", "b" });
                    case "BINARY": return new HostRegistryValue(HostRegistryValueKind.Binary, new byte[] { 1, 255 });
                    default: return null;
                }
            }
        }

        /// <summary>
        /// WshShell.RegRead through the host's IHostRegistryHostService (issue #122): the types and errors of Windows Script Host
        /// (CliParityTests.RegReadReadsTheRegistryAsWsh compares the real registry with cscript); without a host service it is an error
        /// </summary>
        [TestMethod]
        public void RegReadGoesThroughTheHostService()
        {
            var registry = new TestRegistry();
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostRegistryHostService>(() => registry))),
                CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(
                "Dim sh : Set sh = CreateObject(\"WScript.Shell\")\r\n" +
                "Dim v\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Dword\") : collector.Report VarType(v) & \" \" & v\r\n" +
                "v = sh.RegRead(\"HKEY_LOCAL_MACHINE\\Software\\Test\\Sz\") : collector.Report VarType(v) & \" \" & v\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Expand\") : collector.Report v\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Multi\") : collector.Report VarType(v) & \" \" & Join(v, \",\")\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Binary\") : collector.Report VarType(v) & \" \" & TypeName(v(1)) & \" \" & v(1)\r\n" +
                "collector.Report \"[\" & sh.RegRead(\"HKLM\\Software\\Test\\\") & \"]\"\r\n" +
                "On Error Resume Next\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Qword\") : collector.Report Err.Number : Err.Clear\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Test\\Missing\") : collector.Report Err.Number & \" \" & Err.Source & \" \" & Err.Description : Err.Clear\r\n" +
                "v = sh.RegRead(\"HKLM\\Software\\Nope\\Value\") : collector.Report Err.Number & \" \" & Err.Description : Err.Clear\r\n" +
                "v = sh.RegRead(\"hklm\\Software\\Test\\Sz\") : collector.Report Err.Number & \" \" & Err.Description : Err.Clear\r\n" +
                "v = sh.RegRead(\"HKU\\x\\y\") : collector.Report Err.Number : Err.Clear\r\n");
            CollectionAssert.AreEqual(
                new object[]
                {
                    "3 -60",
                    "8 text",
                    @"%USERPROFILE%\x",
                    "8204 a,b",
                    "8204 Byte 255",
                    "[]",
                    13,
                    "-2147024894 WshShell.RegRead Unable to open registry key \"HKLM\\Software\\Test\\Missing\" for reading.",
                    "-2147024894 Invalid root in registry key \"HKLM\\Software\\Nope\\Value\".",
                    "-2147024893 Invalid root in registry key \"hklm\\Software\\Test\\Sz\".",
                    -2147024893,
                },
                collector.Values.ToArray());

            var noService = Assert.Throws<Exception>(() => Execute("collector.Report sh.RegRead(\"HKLM\\Software\\x\")\r\n", null));
            StringAssert.Contains(noService.ToString(), nameof(IHostRegistryHostService));
        }
    }
}
