#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in MAPI.Session (CDO 1.21 - issue #150): the members that the customer scripts use, sending through the mail host service of
    /// the host. CDO 1.21 is not installed on current Windows, so there is no cscript output to compare with - the object model is the one of
    /// CDO 1.21.
    /// </summary>
    [TestClass]
    public sealed class ScriptonMapiSessionTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        private sealed class TestMailHostService : IHostMailHostService
        {
            public List<(HostMailMessage Message, HostSmtpSettings? Smtp)> Sent { get; } = new List<(HostMailMessage, HostSmtpSettings?)>();

            public void Send(HostMailMessage message, HostSmtpSettings? smtp) => Sent.Add((message, smtp));
        }

        private const string CustomerScript =
            "Dim CDOSession, sProfile, oFolder, oMessages, oMsg, oRcpt\r\n" +
            "Dim sServer : sServer = \"exchange01\"\r\n" +
            "Dim sMailbox : sMailbox = \"servicedesk@example.com\"\r\n" +
            "Dim Email_SQL : Email_SQL = \"customer@example.com\"\r\n" +
            "Set CDOSession = CreateObject(\"MAPI.Session\")\r\n" +
            "sProfile = sServer & vbLf & sMailbox\r\n" +
            "CDOSession.logon , , False, True, , False, sProfile\r\n" +
            "Set oFolder = CDOSession.Outbox\r\n" +
            "Set oMessages = oFolder.Messages\r\n" +
            "Set oMsg = oMessages.Add\r\n" +
            "Set oRcpt = oMsg.Recipients\r\n" +
            "oRcpt.Add Email_SQL\r\n" +
            "oRcpt.Resolve\r\n" +
            "oMsg.Subject = \"+++ Problem wurde gelöst +++\"\r\n" +
            "oMsg.Text = \"Ihr Problem wurde gelöst.\"\r\n" +
            "oMsg.Send\r\n" +
            "CDOSession.Logoff\r\n";

        private IScriptControl Engine(TestMailHostService? mail)
        {
            return CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s =>
                {
                    if (mail != null)
                        s.RegisterHostService<IHostMailHostService>(() => mail);
                })),
                CreateScriptControlConfiguration(false, [], []));
        }

        [TestMethod]
        public void TheCustomerScriptSendsOneMailThroughTheMailHostService()
        {
            var mail = new TestMailHostService();
            Engine(mail).ExecuteStatement(CustomerScript);

            (HostMailMessage message, HostSmtpSettings? smtp) = mail.Sent.Single();
            Assert.AreEqual("servicedesk@example.com", message.From, "the mailbox of the ProfileInfo");
            CollectionAssert.AreEqual(new[] { "customer@example.com" }, message.To.ToArray());
            Assert.IsEmpty(message.Cc);
            Assert.IsEmpty(message.Bcc);
            Assert.AreEqual("+++ Problem wurde gelöst +++", message.Subject);
            Assert.AreEqual("Ihr Problem wurde gelöst.", message.TextBody);
            Assert.IsNull(message.HtmlBody);
            Assert.IsNull(smtp, "MAPI.Session has no SMTP settings - the host decides how to send");
        }

        [TestMethod]
        public void RecipientsOfTheTypesCcAndBccAndAddresses()
        {
            var mail = new TestMailHostService();
            var collector = new ValueCollector();
            IScriptControl engine = Engine(mail);
            engine.AddObject("collector", collector, false);
            engine.ExecuteStatement(
                "Dim s : Set s = CreateObject(\"MAPI.Session\")\r\n" +
                "s.Logon , , False, True, , False, \"server\" & vbLf & \"me@example.com\"\r\n" +
                "Dim m : Set m = s.Outbox.Messages.Add(\"Subject\", \"Text\")\r\n" +
                "m.Recipients.Add \"To Name\", \"SMTP:to@example.com\"\r\n" +
                "m.Recipients.Add \"cc@example.com\", , 2\r\n" +
                "m.Recipients.Add \"bcc@example.com\", , 3\r\n" +
                "collector.Report m.Recipients.Count\r\n" +
                "m.Send\r\n" +
                "s.Logoff\r\n");

            Assert.AreEqual(3, collector.Values.Single());
            HostMailMessage message = mail.Sent.Single().Message;
            Assert.AreEqual("Subject", message.Subject);
            Assert.AreEqual("Text", message.TextBody);
            CollectionAssert.AreEqual(new[] { "to@example.com" }, message.To.ToArray());
            CollectionAssert.AreEqual(new[] { "cc@example.com" }, message.Cc.ToArray());
            CollectionAssert.AreEqual(new[] { "bcc@example.com" }, message.Bcc.ToArray());
        }

        [TestMethod]
        public void SendWithoutARecipientRaisesMapiNoRecipients()
        {
            var mail = new TestMailHostService();
            var collector = new ValueCollector();
            IScriptControl engine = Engine(mail);
            engine.AddObject("collector", collector, false);
            engine.ExecuteStatement(
                "On Error Resume Next\r\n" +
                "Dim s : Set s = CreateObject(\"MAPI.Session\")\r\n" +
                "s.Logon , , False, True, , False, \"server\" & vbLf & \"me@example.com\"\r\n" +
                "Dim m : Set m = s.Outbox.Messages.Add\r\n" +
                "m.Send\r\n" +
                "collector.Report Hex(Err.Number) & \"|\" & Err.Source & \"|\" & Err.Description\r\n");

            Assert.AreEqual("80040607|Collaboration Data Objects|[Collaboration Data Objects - [MAPI_E_NO_RECIPIENTS(80040607)]]", collector.Values.Single());
            Assert.IsEmpty(mail.Sent);
        }

        [TestMethod]
        public void LogonWithoutAMailHostServiceRaisesAnErrorThatNamesTheService()
        {
            var collector = new ValueCollector();
            IScriptControl engine = Engine(mail: null);
            engine.AddObject("collector", collector, false);
            engine.ExecuteStatement(
                "On Error Resume Next\r\n" +
                "Dim s : Set s = CreateObject(\"MAPI.Session\")\r\n" +
                "collector.Report TypeName(s)\r\n" +
                "s.Logon , , False, True, , False, \"server\" & vbLf & \"me@example.com\"\r\n" +
                "collector.Report Hex(Err.Number) & \"|\" & Err.Description\r\n" +
                "Err.Clear\r\n" +
                "Dim f : Set f = s.Outbox\r\n" +
                "collector.Report Hex(Err.Number)\r\n");

            Assert.AreEqual("Session", collector.Values[0]);
            StringAssert.StartsWith((string?)collector.Values[1], "80040111|[Collaboration Data Objects - [MAPI_E_LOGON_FAILED(80040111)]]");
            StringAssert.Contains((string?)collector.Values[1], "IHostMailHostService");
            Assert.AreEqual("80040605", collector.Values[2], "MAPI_E_NOT_INITIALIZED without a logon");
        }
    }
}
