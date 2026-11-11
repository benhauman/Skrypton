using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/46 - with tempEnabled: false no file (errors.log included) may be written
    /// </summary>
    [TestClass]
    public sealed class Issue46NoTempFilesWhenDisabledTests : TestBase
    {
        private sealed class RecordingConfiguration : ScriptControlConfiguration
        {
            public RecordingConfiguration(bool tempEnabled, string tempDirectoryPath)
                : base(tempEnabled, tempDirectoryPath, enabledLoadFromDisk: false, translationSuppression: [], noWarn: [])
            {
            }

            public List<string> Writes { get; } = new List<string>();

            protected override void TempFileWriteAllBytes(string folderName, string fileName, byte[] bytes) { Writes.Add(fileName); base.TempFileWriteAllBytes(folderName, fileName, bytes); }
            protected override void TempFileWriteAllLines(string folderName, string fileName, IEnumerable<string> contents) { Writes.Add(fileName); base.TempFileWriteAllLines(folderName, fileName, contents); }
            protected override void TempFileWriteAllText(string folderName, string fileName, string contents) { Writes.Add(fileName); base.TempFileWriteAllText(folderName, fileName, contents); }
        }

        public sealed class ValueCollector
        {
            public object Last { get; private set; }
            public void Report(object value) => Last = value;
        }

        private static string RunAndGetWrites(TestBase test, bool tempEnabled, string script)
        {
            var config = new RecordingConfiguration(tempEnabled, test.TestRunResultsDirectory);
            var collector = new ValueCollector();
            IScriptControl scriptengine = new ScriptControlClass(test.CreateRuntimeHost(test.CreateTestHostServices()), new TestRuntimeLoggerForIssue46(), CultureInfo.InvariantCulture, config);
            scriptengine.Language = "VBScript";
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(script);
            Assert.AreEqual(1, System.Convert.ToInt32(collector.Last));
            return string.Join(",", config.Writes);
        }

        [TestMethod]
        public void NothingIsWrittenWhenTempIsDisabled()
        {
            // A unique script, so that the program is compiled here and not taken from the compiled program cache
            Assert.AreEqual("", RunAndGetWrites(this, tempEnabled: false, "collector.Report 1 ' issue 46, temp disabled\r\n"));
        }

        [TestMethod]
        public void ErrorsLogIsWrittenWhenTempIsEnabled()
        {
            StringAssert.Contains(RunAndGetWrites(this, tempEnabled: true, "collector.Report 1 ' issue 46, temp enabled\r\n"), "errors.log");
        }

        private sealed class TestRuntimeLoggerForIssue46 : Skrypton.RuntimeSupport.Implementations.IRuntimeLogger
        {
            public void LogException(System.Exception exception)
            {
            }
        }
    }
}
