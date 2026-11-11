using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in Msxml2.ServerXMLHTTP.6.0, sending through an offline IHostHttpClientHostService registered by the host (no network access)
    /// </summary>
    [TestClass]
    public sealed class ScriptonServerXMLHTTP60Tests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private (List<object> Values, TestHttpClientHostService Http) Execute(string vbs)
        {
            var http = new TestHttpClientHostService();
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostHttpClientHostService>(() => http))),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(
                "Dim xmlhttp : Set xmlhttp = CreateObject(\"Msxml2.ServerXMLHTTP.6.0\")\r\n" + vbs);
            return (collector.Values, http);
        }

        [TestMethod]
        public void RequestIsSentThroughTheHostService()
        {
            var (values, http) = Execute(
                "xmlhttp.open \"POST\", \"https://example.invalid/api\", False, \"usr\", \"pwd\"\r\n" +
                "xmlhttp.setRequestHeader \"Content-Type\", \"application/json\"\r\n" +
                "xmlhttp.send \"{\"\"a\"\":1}\"\r\n" +
                "collector.Report xmlhttp.readyState & \"|\" & xmlhttp.status & \"|\" & xmlhttp.statusText\r\n");

            Assert.AreEqual("4|200|OK", values[0]);
            HostHttpRequest request = http.Requests.Single();
            Assert.AreEqual("POST", request.Method);
            Assert.AreEqual(new Uri("https://example.invalid/api"), request.Url);
            Assert.AreEqual("usr", request.User);
            Assert.AreEqual("pwd", request.Password);
            Assert.AreEqual("application/json", request.Headers.Single(h => h.Key == "Content-Type").Value);
            Assert.AreEqual("{\"a\":1}", Encoding.UTF8.GetString(request.Body!));
        }

        [TestMethod]
        public void ResponseTextHeadersAndXml()
        {
            var (values, _) = Execute(
                "xmlhttp.open \"GET\", \"https://example.invalid/data\", False\r\n" +
                "xmlhttp.send\r\n" +
                "collector.Report xmlhttp.getResponseHeader(\"content-type\")\r\n" +
                "collector.Report InStr(xmlhttp.responseText, \"<table>\") > 0\r\n" +
                "Dim doc : Set doc = CreateObject(\"Msxml2.DOMDocument\")\r\n" +
                "collector.Report doc.load(xmlhttp.responseXML)\r\n" +
                "collector.Report doc.selectSingleNode(\"//header/name\").text\r\n");

            CollectionAssert.AreEqual(new object[] { "text/xml; charset=utf-8", true, true, "name" }, values.ToArray());
        }

        [TestMethod]
        public void ResponseBeforeSendRaisesAnError()
        {
            var exception = Assert.Throws<Exception>(() => Execute("collector.Report xmlhttp.status\r\n"));
            StringAssert.Contains(exception.ToString(), "only available after send");
        }
    }
}
