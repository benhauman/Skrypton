using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// The built-in Msxml2.DOMDocument: load (file path / URL, Byte array, another DOMDocument), loadXML and xml
    /// </summary>
    [TestClass]
    public sealed class ScriptonMsxml2DOMDocumentTests : TestBase
    {
        public sealed class Host
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
            public byte[] Bytes { get; set; } = Array.Empty<byte>();
            public string Path { get; set; } = "";
        }

        private const string Xml = "<root><item id=\"1\">one</item><item id=\"2\">two</item></root>";

        private List<object> Execute(string vbs, Host host)
        {
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("host", host, false);
            scriptengine.ExecuteStatement(
                "Dim doc : Set doc = CreateObject(\"Msxml2.DOMDocument\")\r\n" +
                "doc.async = False\r\n" +
                vbs);
            return host.Values;
        }

        private const string ReportSecondItem = "host.Report doc.selectSingleNode(\"//item[@id='2']\").text\r\n";

        [TestMethod]
        public void LoadXml()
        {
            var values = Execute("host.Report doc.loadXML(\"" + Xml.Replace("\"", "\"\"") + "\")\r\n" + ReportSecondItem, new Host());
            CollectionAssert.AreEqual(new object[] { true, "two" }, values.ToArray());
        }

        [TestMethod]
        public void LoadXmlReturnsFalseForInvalidXml()
        {
            var values = Execute("host.Report doc.loadXML(\"<root><unclosed></root>\")\r\nhost.Report Len(doc.parseErrorReason) > 0\r\n", new Host());
            CollectionAssert.AreEqual(new object[] { false, true }, values.ToArray());
        }

        [TestMethod]
        public void LoadFromAByteArray()
        {
            var values = Execute("host.Report doc.load(host.Bytes)\r\n" + ReportSecondItem, new Host { Bytes = Encoding.UTF8.GetBytes(Xml) });
            CollectionAssert.AreEqual(new object[] { true, "two" }, values.ToArray());
        }

        [TestMethod]
        public void LoadFromAFilePath()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "skrypton-dom-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, Xml);
            try
            {
                var values = Execute("host.Report doc.load(host.Path)\r\n" + ReportSecondItem, new Host { Path = path });
                CollectionAssert.AreEqual(new object[] { true, "two" }, values.ToArray());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void LoadReturnsFalseForAMissingFile()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "skrypton-dom-missing-" + Guid.NewGuid().ToString("N") + ".xml");
            var values = Execute("host.Report doc.load(host.Path)\r\n", new Host { Path = path });
            CollectionAssert.AreEqual(new object[] { false }, values.ToArray());
        }

        [TestMethod]
        public void LoadFromAnotherDocumentAndXml()
        {
            var values = Execute(
                "Dim other : Set other = CreateObject(\"Msxml2.DOMDocument\")\r\n" +
                "other.loadXML \"" + Xml.Replace("\"", "\"\"") + "\"\r\n" +
                "host.Report doc.load(other)\r\n" +
                "host.Report doc.xml\r\n",
                new Host());
            CollectionAssert.AreEqual(new object[] { true, Xml }, values.ToArray());
        }
    }
}
