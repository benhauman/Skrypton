#nullable enable
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in Shell.Application on the host's IHostProcessControlHostService
    /// </summary>
    [TestClass]
    public sealed class ScriptonShellApplicationTests : TestBase
    {
        private void Execute(string vbs, TestHostProcessControlHostService? processControl)
        {
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (processControl != null)
                        s.RegisterHostService<IHostProcessControlHostService>(() => processControl);
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.ExecuteStatement("Dim sh : Set sh = CreateObject(\"Shell.Application\")\r\n" + vbs);
        }

        [TestMethod]
        public void ShellExecuteGoesThroughTheHostService()
        {
            var processControl = new TestHostProcessControlHostService();
            Execute(
                "sh.ShellExecute \"https://example.invalid\"\r\n" +
                "sh.ShellExecute \"report.pdf\", \"\", \"\", \"print\", 0\r\n",
                processControl);
            CollectionAssert.AreEqual(
                new[] { "shell:https://example.invalid|||1", "shell:report.pdf||print|0" },
                processControl.Started.ToArray());
        }

        [TestMethod]
        public void WithoutAHostServiceShellExecuteRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute("sh.ShellExecute \"https://example.invalid\"\r\n", processControl: null));
            StringAssert.Contains(exception.ToString(), nameof(IHostProcessControlHostService));
        }
    }
}
