#nullable enable
using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Windows.Tests
{
    /// <summary>
    /// The Windows host services (without visible side effects: no windows, no dialogs, the clipboard is restored)
    /// </summary>
    [TestClass]
    public sealed class WindowsHostServicesTests
    {
        [TestMethod]
        public void TheDefaultHostHasTheWindowsServices()
        {
            WindowsRuntimeHost host = WindowsRuntimeHost.CreateDefault();
            Assert.IsInstanceOfType<WindowsHostProcessControlHostService>(host.TryGetRuntimeHostService<IHostProcessControlHostService>());
            Assert.IsInstanceOfType<WindowsHostClipboardHostService>(host.TryGetRuntimeHostService<IHostClipboardHostService>());
            Assert.IsInstanceOfType<WindowsHostMessageBoxHostService>(host.TryGetRuntimeHostService<IHostMessageBoxHostService>());
            Assert.IsInstanceOfType<WindowsHostInputBoxHostService>(host.TryGetRuntimeHostService<IHostInputBoxHostService>());
            Assert.IsInstanceOfType<WindowsHostOleDbConnectionFactoryHostService>(host.TryGetRuntimeHostService<IHostDatabaseConnectionFactoryHostService>());
            Assert.IsInstanceOfType<WindowsHostOutlookHostService>(host.TryGetRuntimeHostService<IHostOutlookHostService>());
            Assert.IsNull(host.TryGetRuntimeHostService<IHostFileSystemHostService>(), "the built-in components use the process file system");
            Assert.IsNull(host.Unregister<IHostOutlookHostService>().TryGetRuntimeHostService<IHostOutlookHostService>());
        }

        [TestMethod]
        public void ExecRunsACommandLineApplication()
        {
            HostProcessExecResult result = new WindowsHostProcessControlHostService().ProcessExec("cmd /c echo hello& exit 2");
            Assert.AreEqual("hello", result.StdOut.Trim());
            Assert.AreEqual(2, result.ExitCode);
        }

        [TestMethod]
        public void RunWaitsForTheExitCodeOfAHiddenProcess()
        {
            Assert.AreEqual(5, new WindowsHostProcessControlHostService().ProcessStart("cmd /c exit 5", 0, true));
        }

        [TestMethod]
        public void TheProcessesIncludeTheTestProcess()
        {
            int id = Process.GetCurrentProcess().Id;
            bool found = false;
            new WindowsHostProcessControlHostService().ProcessesCollect((processId, name) =>
            {
                found |= (processId == id) && name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase);
                return !found;
            });
            Assert.IsTrue(found);
        }

        [TestMethod]
        public void TheClipboardKeepsText()
        {
            var clipboard = new WindowsHostClipboardHostService();
            string? previous = clipboard.GetText("Text");
            try
            {
                Assert.IsTrue(clipboard.SetText("Text", "Skrypton clipboard test"));
                Assert.AreEqual("Skrypton clipboard test", clipboard.GetText("Text"));
                Assert.IsTrue(clipboard.Clear("Text"));
                Assert.IsNull(clipboard.GetText("Text"));
            }
            finally
            {
                if (previous != null)
                    clipboard.SetText("Text", previous);
            }
        }

        [TestMethod]
        public void ADatabaseConnectionWithAnUnknownProviderRaisesAnError()
        {
            Assert.Throws<Exception>(() => new WindowsHostOleDbConnectionFactoryHostService().CreateAndOpenDatabaseConnectionString("Provider=Skrypton.NoSuchProvider;Data Source=x", "", ""), "an unregistered OLE DB provider");
        }
    }
}
