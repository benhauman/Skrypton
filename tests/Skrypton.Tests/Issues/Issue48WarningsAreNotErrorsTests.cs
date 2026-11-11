using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/48 - warnings are errors by default; with treatWarningsAsErrors: false legal VBScript that
    /// produces C# warnings runs with an empty noWarn list
    /// </summary>
    [TestClass]
    public sealed class Issue48WarningsAreNotErrorsTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private sealed class SilentLogger : IRuntimeLogger
        {
            public void LogException(Exception exception)
            {
            }
        }

        private List<object> Execute(string vbs, bool treatWarningsAsErrors)
        {
            // A host configuration: no temp files, no noWarn entries
            var config = new ScriptControlConfiguration(tempEnabled: false, tempDirectoryPath: null, enabledLoadFromDisk: false, translationSuppression: [], noWarn: [], treatWarningsAsErrors: treatWarningsAsErrors);
            var collector = new ValueCollector();
            IScriptControl scriptengine = new ScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), new SilentLogger(), CultureInfo.InvariantCulture, config);
            scriptengine.Language = "VBScript";
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        private const string UnusedVariableAndUnreachableCode =
            "Sub Main\r\n" +
            "  Dim unused : unused = 1\r\n" +
            "  collector.Report \"ran\"\r\n" +
            "  Exit Sub\r\n" +
            "  unused = 2\r\n" +
            "End Sub\r\n" +
            "Main\r\n";

        [TestMethod]
        [DataRow(UnusedVariableAndUnreachableCode, DisplayName = "CS0219 / CS0162: unused variable, code after Exit Sub")]
        [DataRow("Function F()\r\n  F = 1\r\n  Exit Function\r\n  F = 2\r\nEnd Function\r\ncollector.Report \"ran\" & F()\r\n", DisplayName = "CS0162: code after Exit Function")]
        [DataRow("Dim declaredOnly\r\ncollector.Report \"ran\"\r\n", DisplayName = "CS0168: declared but never used")]
        [DataRow("Dim x : x = 1\r\nx = x\r\ncollector.Report \"ran\" & x\r\n", DisplayName = "CS1717: x = x")]
        public void LegalVBScriptRunsWithAnEmptyNoWarnList(string vbs)
        {
            List<object> values = Execute(vbs, treatWarningsAsErrors: false);
            Assert.HasCount(1, values);
            StringAssert.StartsWith((string)values[0], "ran");
        }

        [TestMethod]
        public void WarningsAreErrorsByDefault()
        {
            var config = new ScriptControlConfiguration(tempEnabled: false, tempDirectoryPath: null, enabledLoadFromDisk: false, translationSuppression: [], noWarn: []);
            Assert.IsTrue(config.TreatWarningsAsErrors);
        }

        [TestMethod]
        public void WarningsAsErrorsFailsTheCompilation()
        {
            Assert.ThrowsExactly<CompilationFailedException>(() => Execute(UnusedVariableAndUnreachableCode, treatWarningsAsErrors: true));
        }
    }
}
