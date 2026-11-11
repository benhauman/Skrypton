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
    /// The built-in InternetExplorer.Application (pages are opened through IHostProcessControlHostService once the browser is visible) and htmlfile
    /// (parentWindow.clipboardData goes through IHostClipboardHostService); see ComParityTests for the comparison with Windows
    /// </summary>
    [TestClass]
    public sealed class ScriptonInternetExplorerHtmlFileTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        private List<object?> Execute(string vbs, TestHostProcessControlHostService? processControl = null, TestClipboardHostService? clipboard = null)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (processControl != null)
                        s.RegisterHostService<IHostProcessControlHostService>(() => processControl);
                    if (clipboard != null)
                        s.RegisterHostService<IHostClipboardHostService>(() => clipboard);
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        [TestMethod]
        public void AVisibleBrowserOpensThePageWithTheDefaultBrowser()
        {
            var processControl = new TestHostProcessControlHostService();
            Execute(
                "Dim ie : Set ie = CreateObject(\"InternetExplorer.Application\")\r\n" +
                "ie.Visible = True\r\n" +
                "ie.Navigate \"https://example.invalid/a\"\r\n",
                processControl);
            CollectionAssert.AreEqual(new[] { "shell:https://example.invalid/a||open|1" }, processControl.Started.ToArray());
        }

        [TestMethod]
        public void AnInvisibleBrowserOpensThePageWhenItBecomesVisible()
        {
            const string navigate =
                "Dim ie : Set ie = CreateObject(\"InternetExplorer.Application\")\r\n" +
                "ie.Navigate \"https://example.invalid/b\"\r\n" +
                "collector.Report ie.LocationURL\r\n" +
                "collector.Report ie.ReadyState\r\n";

            var invisible = new TestHostProcessControlHostService();
            var values = Execute(navigate, invisible);
            Assert.AreEqual("https://example.invalid/b", values[0]);
            Assert.AreEqual(4, Convert.ToInt32(values[1]));
            Assert.IsEmpty(invisible.Started, "an invisible browser opens nothing");

            var visible = new TestHostProcessControlHostService();
            Execute(navigate + "ie.Visible = True\r\nie.Visible = True\r\n", visible);
            CollectionAssert.AreEqual(new[] { "shell:https://example.invalid/b||open|1" }, visible.Started.ToArray(), "opened once");
        }

        [TestMethod]
        public void NavigatingAVisibleBrowserWithoutAHostServiceRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute(
                "Dim ie : Set ie = CreateObject(\"InternetExplorer.Application\")\r\n" +
                "ie.Visible = True\r\n" +
                "ie.Navigate \"https://example.invalid\"\r\n"));
            StringAssert.Contains(exception.ToString(), nameof(IHostProcessControlHostService));
        }

        [TestMethod]
        public void ABrowserThatWasQuitCannotBeUsed()
        {
            Assert.Throws<Exception>(() => Execute(
                "Dim ie : Set ie = CreateObject(\"InternetExplorer.Application\")\r\n" +
                "ie.Quit\r\n" +
                "ie.Navigate \"about:blank\"\r\n"));
        }

        [TestMethod]
        public void ClipboardDataGoesThroughTheHostService()
        {
            var clipboard = new TestClipboardHostService();
            var values = Execute(
                "Dim h : Set h = CreateObject(\"htmlfile\")\r\n" +
                "Dim cd : Set cd = h.parentWindow.clipboardData\r\n" +
                "collector.Report cd.setData(\"text\", \"hello\")\r\n" +
                "collector.Report cd.getData(\"Text\")\r\n" +
                "cd.clearData \"Text\"\r\n" +
                "collector.Report cd.getData(\"Text\")\r\n",
                clipboard: clipboard);
            CollectionAssert.AreEqual(new object?[] { true, "hello", "" }, values.ToArray());
        }

        [TestMethod]
        public void ClipboardDataWithoutAHostServiceRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute("Dim h : Set h = CreateObject(\"htmlfile\")\r\nh.parentWindow.clipboardData.setData \"Text\", \"x\"\r\n"));
            StringAssert.Contains(exception.ToString(), nameof(IHostClipboardHostService));
        }
    }
}
