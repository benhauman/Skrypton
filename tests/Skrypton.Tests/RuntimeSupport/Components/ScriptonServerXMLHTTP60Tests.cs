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
            Assert.AreEqual("{\"a\":1}", Encoding.UTF8.GetString(request.Body));
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

        private (List<object> Values, TestHttpClientHostService Http) ExecuteScript(string vbs, Uri responseUrl = null)
        {
            var http = new TestHttpClientHostService { ResponseUrl = responseUrl };
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostHttpClientHostService>(() => http))),
                CreateScriptControlConfiguration(false, [], [])
            );
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return (collector.Values, http);
        }

        /// <summary>
        /// Msxml2.XMLHTTP and its versions are the same built-in (TypeName IServerXMLHTTPRequest2 for all of them with Windows VBScript); send
        /// takes the xml of a DOMDocument or the document itself (issue #102)
        /// </summary>
        [TestMethod]
        [DataRow("MSXML2.XMLHTTP")]
        [DataRow("Msxml2.XMLHTTP.3.0")]
        [DataRow("Msxml2.XMLHTTP.6.0")]
        [DataRow("Microsoft.XMLHTTP")]
        [DataRow("Msxml2.ServerXMLHTTP")]
        public void XmlHttpProgIdsAreTheServerXmlHttpBuiltIn(string progId)
        {
            var (values, http) = ExecuteScript(
                "Dim objDom : Set objDom = CreateObject(\"MSXML2.DOMDocument\")\r\n" +
                "objDom.async = False\r\n" +
                "objDom.loadXML \"<Predict>printer</Predict>\"\r\n" +
                "Dim x : Set x = CreateObject(\"" + progId + "\")\r\n" +
                "x.open \"POST\", \"http://example.invalid/predict.asmx\", False\r\n" +
                "x.setRequestHeader \"Content-Type\", \"text/xml; charset=utf-8\"\r\n" +
                "x.setRequestHeader \"SOAPAction\", \"urn:predict\"\r\n" +
                "x.send objDom.xml\r\n" +
                "collector.Report TypeName(x) & \"|\" & x.status & \"|\" & (Len(x.responseText) > 0)\r\n" +
                "x.open \"POST\", \"http://example.invalid/doc\", False\r\n" +
                "x.send objDom\r\n");
            Assert.AreEqual("IServerXMLHTTPRequest2|200|True", values.Single());
            Assert.HasCount(2, http.Requests);
            foreach (HostHttpRequest request in http.Requests)
                Assert.AreEqual("<Predict>printer</Predict>\r\n", Encoding.UTF8.GetString(request.Body)); // MSXML's xml ends with CRLF
            Assert.AreEqual("urn:predict", http.Requests[0].Headers.Single(h => h.Key == "SOAPAction").Value);
        }

        /// <summary>
        /// WinHttp.WinHttpRequest.5.1 (issue #102): on the same transport; Option(1) is the URL of the response (after redirects), other options
        /// are kept; WinHttp's User-Agent and, for a text body, Content-Type "text/plain; Charset=UTF-8" unless the script sets them (as
        /// received from Windows' WinHttp)
        /// </summary>
        [TestMethod]
        public void WinHttpRequest()
        {
            var (values, http) = ExecuteScript(
                "Dim w : Set w = CreateObject(\"WinHttp.WinHttpRequest.5.1\")\r\n" +
                "w.setTimeouts 59000, 59000, 59000, 59000\r\n" +
                "w.Open \"GET\", \"http://example.invalid/sap/ts_interface\", False\r\n" +
                "collector.Report w.Option(1)\r\n" +
                "w.setRequestHeader \"User-Agent\", \"Mozilla/5.0\"\r\n" +
                "w.setRequestHeader \"Authorization\", \"Basic dXNlcjpwd2Q=\"\r\n" +
                "w.Send\r\n" +
                "collector.Report TypeName(w) & \"|\" & w.Status & \"|\" & w.StatusText & \"|\" & w.Option(1) & \"|\" & (Len(w.ResponseText) > 0) & \"|\" & w.GetResponseHeader(\"Content-Type\")\r\n" +
                "w.Option(4) = 13056\r\n" +
                "collector.Report w.Option(4) & \"|\" & w.Option(0)\r\n" +
                "w.Open \"POST\", \"http://example.invalid/form\", False\r\n" +
                "w.Send \"a=1\"\r\n",
                responseUrl: new Uri("http://example.invalid/redirected"));
            CollectionAssert.AreEqual(
                new object[]
                {
                    "http://example.invalid/sap/ts_interface",
                    "WinHttpRequest|200|OK|http://example.invalid/redirected|True|text/xml; charset=utf-8",
                    "13056|Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)",
                },
                values.ToArray());
            HostHttpRequest get = http.Requests[0];
            Assert.AreEqual(TimeSpan.FromMilliseconds(59000), get.Timeout);
            Assert.AreEqual("Mozilla/5.0", get.Headers.Single(h => h.Key == "User-Agent").Value);
            Assert.AreEqual("Basic dXNlcjpwd2Q=", get.Headers.Single(h => h.Key == "Authorization").Value);
            HostHttpRequest post = http.Requests[1];
            Assert.AreEqual("a=1", Encoding.UTF8.GetString(post.Body));
            Assert.AreEqual("text/plain; Charset=UTF-8", post.Headers.Single(h => h.Key == "Content-Type").Value);
            Assert.AreEqual("Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)", post.Headers.Single(h => h.Key == "User-Agent").Value);
        }
    }
}
