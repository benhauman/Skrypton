#nullable enable
using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in Outlook.Application: saving, sending, displaying, finding items and resolving recipients go through IHostOutlookHostService
    /// (see ComParityTests for the comparison of the object model with Outlook)
    /// </summary>
    [TestClass]
    public sealed class ScriptonOutlookApplicationTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        private List<object?> Execute(string vbs, TestOutlookHostService? outlook)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (outlook != null)
                        s.RegisterHostService<IHostOutlookHostService>(() => outlook);
                })),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("Dim o : Set o = CreateObject(\"Outlook.Application\")\r\n" + vbs);
            return collector.Values;
        }

        [TestMethod]
        public void AMailIsSentWithItsPropertiesAndRecipients()
        {
            var outlook = new TestOutlookHostService();
            var values = Execute(
                "Dim m : Set m = o.CreateItem(0)\r\n" +
                "m.Subject = \"Report\" : m.To = \"a@example.invalid\" : m.CC = \"b@example.invalid\" : m.HTMLBody = \"<p>x</p>\"\r\n" +
                "m.Recipients.Add \"c@example.invalid\"\r\n" +
                "m.Send\r\n" +
                "collector.Report m.Sent\r\n",
                outlook);
            Assert.AreEqual(true, values[0]);
            HostOutlookItem sent = outlook.Sent[0];
            Assert.AreEqual(HostOutlookItemType.MailItem, sent.ItemType);
            Assert.AreEqual("Report", sent.GetProperty("Subject"));
            Assert.AreEqual("a@example.invalid; c@example.invalid", sent.GetProperty("To"));
            Assert.AreEqual("b@example.invalid", sent.GetProperty("CC"));
            Assert.AreEqual("<p>x</p>", sent.GetProperty("HTMLBody"));
        }

        [TestMethod]
        public void AMailWithoutRecipientsCannotBeSent()
        {
            Assert.Throws<Exception>(() => Execute("Dim m : Set m = o.CreateItem(0)\r\nm.Send\r\n", new TestOutlookHostService()));
        }

        [TestMethod]
        public void SavedContactsAreFoundInTheContactsFolder()
        {
            var outlook = new TestOutlookHostService();
            var values = Execute(
                "Dim c : Set c = o.CreateItem(2)\r\n" +
                "c.FirstName = \"Max\" : c.LastName = \"Muster\" : c.Save\r\n" +
                "collector.Report c.Saved\r\n" +
                "collector.Report c.EntryID\r\n" +
                "Dim f : Set f = o.GetNamespace(\"MAPI\").GetDefaultFolder(10)\r\n" +
                "collector.Report f.Items.Count\r\n" +
                "collector.Report f.Items.Find(\"[LastName] = 'muster' And [FirstName] = \"\"Max\"\"\").FullName\r\n" +
                "collector.Report TypeName(f.Items.Find(\"[LastName] = 'Other'\"))\r\n" +
                "collector.Report f.Items(1).FileAs\r\n",
                outlook);
            CollectionAssert.AreEqual(new object?[] { true, "ENTRY1", 1, "Max Muster", "Nothing", "Muster, Max" }, values.ToArray());
        }

        [TestMethod]
        public void ItemsAreDisplayedAndRecipientsResolvedThroughTheHost()
        {
            var outlook = new TestOutlookHostService();
            var values = Execute(
                "Dim a : Set a = o.GetNamespace(\"MAPI\").GetDefaultFolder(9).Items.Add()\r\n" +
                "a.Subject = \"Termin\" : a.Display True\r\n" +
                "Dim r : Set r = o.GetNamespace(\"MAPI\").CreateRecipient(\"x@example.invalid\")\r\n" +
                "collector.Report r.Resolve()\r\n" +
                "collector.Report o.GetNamespace(\"MAPI\").CreateRecipient(\"nobody\").Resolve()\r\n",
                outlook);
            CollectionAssert.AreEqual(new object?[] { true, false }, values.ToArray());
            Assert.AreEqual(HostOutlookItemType.AppointmentItem, outlook.Displayed[0].Item.ItemType);
            Assert.IsTrue(outlook.Displayed[0].Modal);
        }

        [TestMethod]
        public void WithoutAHostServiceItemsCanBeFilledInButNotSaved()
        {
            var values = Execute("Dim m : Set m = o.CreateItem(0)\r\nm.Subject = \"x\"\r\ncollector.Report m.Subject\r\n", outlook: null);
            Assert.AreEqual("x", values[0]);
            var exception = Assert.Throws<Exception>(() => Execute("Dim c : Set c = o.CreateItem(2)\r\nc.Save\r\n", outlook: null));
            StringAssert.Contains(exception.ToString(), nameof(IHostOutlookHostService));
        }
    }
}
