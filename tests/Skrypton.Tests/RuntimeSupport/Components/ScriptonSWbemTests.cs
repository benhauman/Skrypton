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
            CollectionAssert.AreEqual(new object[] { "666:xxx;4444:IExplore.exe;", "SWbemServicesEx" }, values.ToArray());
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

            var filtered = Assert.Throws<Exception>(() => Execute("GetObject(\"winmgmts:\").ExecQuery \"SELECT * FROM Win32_TimeZone WHERE Bias = '60'\"\r\n"));
            StringAssert.Contains(filtered.ToString(), "Win32_TimeZone");
        }

        /// <summary>
        /// Win32_TimeZone is the local time zone (issue #101) - no host service needed; Bias is the offset from UTC in minutes (60 for UTC+01:00,
        /// as with Windows WMI - CliParityTests.WmiTimeZoneIsTheOneOfWindows compares it with cscript)
        /// </summary>
        [TestMethod]
        public void Win32TimeZoneIsTheLocalTimeZone()
        {
            var values = Execute(
                "Dim svc : Set svc = GetObject(\"winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2\")\r\n" +
                "Dim col : Set col = svc.ExecQuery(\"Select * from Win32_TimeZone\")\r\n" +
                "collector.Report col.Count\r\n" +
                "Dim tz, n\r\n" +
                "For Each tz In col\r\n" +
                "    n = n + 1\r\n" +
                "    collector.Report tz.Bias\r\n" +
                "    collector.Report VarType(tz.Bias) & \"|\" & VarType(tz.DaylightDayOfWeek) & \"|\" & VarType(tz.SettingID) & \"|\" & TypeName(tz)\r\n" +
                "    collector.Report tz.Caption & \"|\" & tz.StandardName & \"|\" & tz.DaylightName\r\n" +
                "    Exit For\r\n" +
                "Next\r\n" +
                "collector.Report n\r\n" +
                "collector.Report svc.InstancesOf(\"Win32_TimeZone\").Count\r\n",
                registerProcessControl: false);
            TimeZoneInfo local = TimeZoneInfo.Local;
            CollectionAssert.AreEqual(
                new object[] { 1, (int)local.BaseUtcOffset.TotalMinutes, "3|17|1|SWbemObjectEx", local.DisplayName + "|" + local.StandardName + "|" + local.DaylightName, (short)1, 1 },
                values.ToArray());
        }
    }
}
