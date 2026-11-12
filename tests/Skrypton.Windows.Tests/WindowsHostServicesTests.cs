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
            Assert.IsInstanceOfType<WindowsHostEnvironmentVariablesHostService>(host.TryGetRuntimeHostService<IHostEnvironmentVariablesHostService>());
            Assert.IsNull(host.TryGetRuntimeHostService<IHostFileSystemHostService>(), "the built-in components use the process file system");
            Assert.IsNull(host.Unregister<IHostOutlookHostService>().TryGetRuntimeHostService<IHostOutlookHostService>());
        }

        /// <summary>
        /// ADODB inside a host's TransactionScope (issue #99): the connection does not enlist in it, as with Windows ADO - the text provider
        /// (no local transactions) would fail at Open otherwise. The query runs and the ambient transaction stays a local one.
        /// </summary>
        [TestMethod]
        public void AdodbConnectionDoesNotEnlistInTheAmbientTransaction()
        {
            string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "skrypton-ado-" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(folder);
            try
            {
                System.IO.File.WriteAllText(System.IO.Path.Combine(folder, "products.csv"), "productid,equinr\r\n4711,E1\r\n4712,E2\r\n4714,E1\r\n");
                string script =
                    "Dim cn : Set cn = CreateObject(\"ADODB.Connection\")\r\n" +
                    "cn.ConnectionString = \"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + folder + ";Extended Properties=\"\"text;HDR=Yes;FMT=Delimited\"\"\"\r\n" +
                    "cn.Open\r\n" +
                    "Dim rs : Set rs = cn.Execute(\"SELECT productid FROM [products.csv] WHERE equinr = 'E1'\")\r\n" +
                    "Do While Not rs.EOF\r\n" +
                    "    WScript.Echo rs.Fields(0).Value\r\n" +
                    "    rs.MoveNext\r\n" +
                    "Loop\r\n" +
                    "rs.Close\r\n" +
                    "cn.Close\r\n" +
                    "WScript.Echo \"State=\" & cn.State\r\n";
                using (new System.Transactions.TransactionScope())
                {
                    var output = new System.IO.StringWriter();
                    var error = new System.IO.StringWriter();
                    int exitCode = new ScriptonScriptRunner(WindowsRuntimeHost.CreateDefault(), output, error, System.IO.TextReader.Null, System.Globalization.CultureInfo.GetCultureInfo("en-US")).Run(script, System.IO.Path.Combine(folder, "repro.vbs"), Array.Empty<string>());
                    if (error.ToString().Contains("provider is not registered"))
                        Assert.Inconclusive("The Microsoft.ACE.OLEDB.12.0 provider is not installed: " + error);
                    Assert.AreEqual("", error.ToString());
                    Assert.AreEqual(0, exitCode);
                    CollectionAssert.AreEqual(new[] { "4711", "4714", "State=0" }, output.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                    Assert.AreEqual(Guid.Empty, System.Transactions.Transaction.Current!.TransactionInformation.DistributedIdentifier);
                }
            }
            finally
            {
                System.IO.Directory.Delete(folder, recursive: true);
            }
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
