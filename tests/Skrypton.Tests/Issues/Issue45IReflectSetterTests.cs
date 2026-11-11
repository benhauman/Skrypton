using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Compat;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/45 - assigning a property through IReflect must find instance properties
    /// </summary>
    [TestClass]
    public sealed class Issue45IReflectSetterTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        public sealed class Dict : BasicIReflectImplementation
        {
            private readonly Dictionary<string, object> _items = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            public int CompareMode { get; set; }

            [IndexerName("Item")]
            public object this[string key]
            {
                get => _items[key];
                set => _items[key] = value;
            }
        }

        [TestMethod]
        public void InstanceAndIndexedPropertiesCanBeSetAndReadBack()
        {
            var collector = new ValueCollector();
            var dict = new Dict();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.AddObject("d", dict, false);
            scriptengine.ExecuteStatement(
                "d.CompareMode = 1\r\n" +
                "d.Item(\"a\") = \"x\"\r\n" +
                "collector.Report d.CompareMode\r\n" +
                "collector.Report d.Item(\"a\")\r\n");

            Assert.AreEqual(1, dict.CompareMode);
            Assert.AreEqual("x", dict["a"]);
            Assert.HasCount(2, collector.Values);
            Assert.AreEqual(1, Convert.ToInt32(collector.Values[0]));
            Assert.AreEqual("x", collector.Values[1]);
        }
    }
}
