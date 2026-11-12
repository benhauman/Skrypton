#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in CDO.Message / CDO.Configuration (issue #149): sent over SMTP to a loopback server (the default mail host service) - the
    /// envelope, AUTH LOGIN, the headers and the bodies as Windows CDOSYS sends them to the same server - and through a mail host service of the host
    /// </summary>
    [TestClass]
    public sealed class ScriptonCdoMessageTests : TestBase
    {
        private const string SendScript =
            "Dim oMsg, oCfg, oCfgFields\r\n" +
            "Const cdo = \"http://schemas.microsoft.com/cdo/configuration/\"\r\n" +
            "Set oCfg = CreateObject(\"CDO.Configuration\")\r\n" +
            "Set oCfgFields = oCfg.Fields\r\n" +
            "oCfgFields(cdo & \"languagecode\") = \"de\"\r\n" +
            "oCfgFields(cdo & \"sendusing\") = 2\r\n" +
            "oCfgFields(cdo & \"smtpserverport\") = PORT\r\n" +
            "oCfgFields(cdo & \"smtpserver\") = \"127.0.0.1\"\r\n" +
            "oCfgFields(cdo & \"smtpauthenticate\") = 1\r\n" +
            "oCfgFields(cdo & \"sendusername\") = \"user1\"\r\n" +
            "oCfgFields(cdo & \"sendpassword\") = \"secret\"\r\n" +
            "oCfgFields.update\r\n" +
            "Set oMsg = CreateObject(\"CDO.Message\")\r\n" +
            "Set oMsg.Configuration = oCfg\r\n" +
            "oMsg.Subject = \"Incident 4711 solved\"\r\n" +
            "oMsg.From = \"Service Desk <ServiceDesk@example.com>\"\r\n" +
            "oMsg.To = \"agent@example.com\"\r\n" +
            "oMsg.CC = \"Lead <lead@example.com>\"\r\n" +
            "oMsg.BCC = \"audit@example.com\"\r\n" +
            "oMsg.HTMLBody = \"<p>The incident <b>4711</b> is solved.</p>\"\r\n" +
            "oMsg.Send\r\n";

        /// <summary>
        /// CDOSYS (cscript) sent this script to the same kind of server: AUTH LOGIN with user1 / secret, MAIL FROM ServiceDesk@example.com, RCPT TO
        /// for To, CC and BCC, the headers From "Service Desk", To, Cc (no Bcc) and Subject, and multipart/alternative with the text of the HTML
        /// </summary>
        [TestMethod]
        public void SendsOverSmtpAsCdosysDoes()
        {
            using var server = new LoopbackSmtpServer();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices(s => { })), CreateScriptControlConfiguration(false, [], []));
            scriptengine.ExecuteStatement(SendScript.Replace("PORT", server.Port.ToString(System.Globalization.CultureInfo.InvariantCulture)));

            string[] commands = server.Commands();
            // AUTH LOGIN - CDOSYS sends the user name on a line of its own, SmtpClient with the command ("AUTH login <user>"); the same for SMTP
            string[] login = string.Join(" ", commands.SkipWhile(c => !c.StartsWith("AUTH LOGIN", StringComparison.OrdinalIgnoreCase)).Take(3))
                .Split(' ').Skip(2).Take(2).ToArray();
            CollectionAssert.AreEqual(
                new[] { Convert.ToBase64String(Encoding.ASCII.GetBytes("user1")), Convert.ToBase64String(Encoding.ASCII.GetBytes("secret")) },
                login, string.Join(Environment.NewLine, commands));
            Assert.AreEqual("MAIL FROM:<ServiceDesk@example.com>", commands.Single(c => c.StartsWith("MAIL FROM", StringComparison.OrdinalIgnoreCase)).Replace(": <", ":<")); // CDOSYS writes "MAIL FROM: <..>"
            // the recipients of To, CC and BCC - CDOSYS sends them in this order, SmtpClient To, BCC, CC
            CollectionAssert.AreEquivalent(
                new[] { "RCPT TO:<agent@example.com>", "RCPT TO:<lead@example.com>", "RCPT TO:<audit@example.com>" },
                commands.Where(c => c.StartsWith("RCPT TO", StringComparison.OrdinalIgnoreCase)).Select(c => c.Replace(": <", ":<")).ToArray());

            string data = server.Data();
            StringAssert.Contains(data, "From: \"Service Desk\" <ServiceDesk@example.com>");
            StringAssert.Matches(data, new System.Text.RegularExpressions.Regex(@"(?m)^To: <?agent@example\.com>?\r?$"));
            StringAssert.Matches(data, new System.Text.RegularExpressions.Regex(@"(?m)^Cc: ""?Lead""? <lead@example\.com>\r?$"));
            Assert.DoesNotContain("audit@example.com", data, "BCC is not a header");
            StringAssert.Contains(data, "Subject: Incident 4711 solved");
            StringAssert.Contains(data, "multipart/alternative");
            StringAssert.Contains(data, "text/plain");
            StringAssert.Contains(data, "The incident 4711 is solved.");
            StringAssert.Contains(data, "text/html");
            StringAssert.Contains(data, "<p>The incident <b>4711</b> is solved.</p>");
        }

        [TestMethod]
        public void SendsThroughTheMailHostServiceOfTheHost()
        {
            var mail = new TestMailHostService();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostMailHostService>(() => mail))), CreateScriptControlConfiguration(false, [], []));
            scriptengine.ExecuteStatement(SendScript.Replace("PORT", "2525"));

            (HostMailMessage message, HostSmtpSettings? smtp) = mail.Sent.Single();
            Assert.AreEqual("Service Desk <ServiceDesk@example.com>", message.From);
            CollectionAssert.AreEqual(new[] { "agent@example.com" }, message.To.ToArray());
            CollectionAssert.AreEqual(new[] { "Lead <lead@example.com>" }, message.Cc.ToArray());
            CollectionAssert.AreEqual(new[] { "audit@example.com" }, message.Bcc.ToArray());
            Assert.AreEqual("Incident 4711 solved", message.Subject);
            Assert.AreEqual("<p>The incident <b>4711</b> is solved.</p>", message.HtmlBody);
            Assert.AreEqual("The incident 4711 is solved.\r\n\r\n", message.TextBody); // as CDOSYS makes it
            Assert.IsNotNull(smtp);
            Assert.AreEqual("127.0.0.1", smtp.Server);
            Assert.AreEqual(2525, smtp.Port);
            Assert.AreEqual("user1", smtp.UserName);
            Assert.AreEqual("secret", smtp.Password);
            Assert.IsFalse(smtp.UseSsl);
            Assert.AreEqual(TimeSpan.FromSeconds(30), smtp.Timeout);
        }

        [TestMethod]
        public void SeveralAddressesAreSeparatedByCommaOrSemicolon()
        {
            var mail = new TestMailHostService();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostMailHostService>(() => mail))), CreateScriptControlConfiguration(false, [], []));
            scriptengine.ExecuteStatement(SendScript.Replace("PORT", "25").Replace(
                "oMsg.To = \"agent@example.com\"", "oMsg.To = \"a@example.com; \"\"Doe, John\"\" <john@example.com>, Lead <lead@example.com>;\""));

            CollectionAssert.AreEqual(
                new[] { "a@example.com", "\"Doe, John\" <john@example.com>", "Lead <lead@example.com>" },
                mail.Sent.Single().Message.To.ToArray());
        }

        private sealed class TestMailHostService : IHostMailHostService
        {
            public List<(HostMailMessage Message, HostSmtpSettings? Smtp)> Sent { get; } = new List<(HostMailMessage, HostSmtpSettings?)>();

            public void Send(HostMailMessage message, HostSmtpSettings? smtp) => Sent.Add((message, smtp));
        }

        /// <summary>
        /// An SMTP server on a loopback port for one session: it offers AUTH LOGIN, accepts everything and keeps the commands and the data
        /// </summary>
        private sealed class LoopbackSmtpServer : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(IPAddress.Loopback, 0);
            private readonly List<string> _commands = new List<string>();
            private readonly StringBuilder _data = new StringBuilder();
            private readonly Thread _thread;

            public LoopbackSmtpServer()
            {
                _listener.Start();
                Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                _thread = new Thread(Serve) { IsBackground = true };
                _thread.Start();
            }

            public int Port { get; }

            public string[] Commands()
            {
                _thread.Join(TimeSpan.FromSeconds(30));
                lock (_commands)
                    return _commands.ToArray();
            }

            public string Data()
            {
                _thread.Join(TimeSpan.FromSeconds(30));
                lock (_commands)
                    return _data.ToString();
            }

            private void Serve()
            {
                using TcpClient client = _listener.AcceptTcpClient();
                using NetworkStream stream = client.GetStream();
                using var reader = new StreamReader(stream, Encoding.ASCII);
                using var writer = new StreamWriter(stream, Encoding.ASCII) { NewLine = "\r\n", AutoFlush = true };
                writer.WriteLine("220 localhost ESMTP test");
                bool data = false;
                int authLines = 0;
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    lock (_commands)
                    {
                        if (data)
                        {
                            if (line == ".")
                            {
                                data = false;
                                writer.WriteLine("250 queued");
                            }
                            else
                            {
                                _data.Append(line).Append("\r\n");
                            }
                            continue;
                        }
                        _commands.Add(line);
                    }
                    if (authLines > 0)
                    {
                        authLines--;
                        writer.WriteLine((authLines > 0) ? "334 UGFzc3dvcmQ6" : "235 ok");
                        continue;
                    }
                    string command = line.ToUpperInvariant();
                    if (command.StartsWith("EHLO", StringComparison.Ordinal))
                    {
                        writer.WriteLine("250-localhost");
                        writer.WriteLine("250 AUTH LOGIN");
                    }
                    else if (command == "AUTH LOGIN")
                    {
                        authLines = 2;
                        writer.WriteLine("334 VXNlcm5hbWU6");
                    }
                    else if (command.StartsWith("AUTH LOGIN ", StringComparison.Ordinal)) // with the user name
                    {
                        authLines = 1;
                        writer.WriteLine("334 UGFzc3dvcmQ6");
                    }
                    else if (command == "DATA")
                    {
                        data = true;
                        writer.WriteLine("354 go");
                    }
                    else if (command == "QUIT")
                    {
                        writer.WriteLine("221 bye");
                        break;
                    }
                    else
                    {
                        writer.WriteLine("250 ok");
                    }
                }
            }

            public void Dispose() => _listener.Stop();
        }
    }
}
