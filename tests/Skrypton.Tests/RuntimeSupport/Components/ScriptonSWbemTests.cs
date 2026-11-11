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
    /// The built-in WMI (GetObject("winmgmts:"), WbemScripting.SWbemLocator) on the host's IHostProcessControlHostService
    /// </summary>
    [TestClass]
    public sealed class ScriptonSWbemTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private List<object> Execute(string vbs, bool registerProcessControl = true)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (registerProcessControl)
                        s.RegisterHostService<IHostProcessControlHostService>(() => new TestHostProcessControlHostService()); // processes 666 "xxx" and 4444 "IExplore.exe"
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        [TestMethod]
        public void InstancesOfWin32Process()
        {
            var values = Execute(
                "Dim p, s\r\n" +
                "For Each p In GetObject(\"winmgmts:\").InstancesOf(\"Win32_Process\")\r\n" +
                "  s = s & p.ProcessId & \":\" & p.Name & \";\"\r\n" +
                "Next\r\n" +
                "collector.Report s\r\n" +
                "collector.Report TypeName(GetObject(\"winmgmts:\"))\r\n");
            CollectionAssert.AreEqual(new object[] { "666:xxx;4444:IExplore.exe;", "SWbemServices" }, values.ToArray());
        }

        [TestMethod]
        public void ExecQueryWithWhereThroughTheLocator()
        {
            var values = Execute(
                "Dim locator : Set locator = CreateObject(\"WbemScripting.SWbemLocator\")\r\n" +
                "Dim wmi : Set wmi = locator.ConnectServer(\".\", \"root\\cimv2\")\r\n" +
                "Dim found : Set found = wmi.ExecQuery(\"SELECT * FROM Win32_Process WHERE Name = 'iexplore.exe'\")\r\n" +
                "collector.Report found.Count\r\n" +
                "collector.Report found.Item(\"Win32_Process.Handle=\"\"4444\"\"\").Name\r\n");
            CollectionAssert.AreEqual(new object[] { 1, "IExplore.exe" }, values.ToArray());
        }

        [TestMethod]
        public void UnsupportedClassAndMissingHostServiceRaiseErrors()
        {
            var unsupported = Assert.Throws<Exception>(() => Execute("GetObject(\"winmgmts:\").InstancesOf \"Win32_Service\"\r\n"));
            StringAssert.Contains(unsupported.ToString(), "Win32_Service");

            var noService = Assert.Throws<Exception>(() => Execute("GetObject(\"winmgmts:\").InstancesOf \"Win32_Process\"\r\n", registerProcessControl: false));
            StringAssert.Contains(noService.ToString(), nameof(IHostProcessControlHostService));
        }
    }
}
