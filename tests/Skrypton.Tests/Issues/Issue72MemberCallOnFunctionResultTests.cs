using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/72 - a member call directly on a CreateObject(...) (or other function) result must translate
    /// </summary>
    [TestClass]
    public sealed class Issue72MemberCallOnFunctionResultTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private List<object> Execute(string vbs)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        [TestMethod]
        public void MemberCallOnACreateObjectResult()
        {
            var values = Execute(
                "collector.Report CreateObject(\"Scripting.Dictionary\").Count\r\n" +
                "Dim e : e = CreateObject(\"Scripting.Dictionary\").Exists(\"a\")\r\n" +
                "collector.Report e\r\n");
            Assert.AreEqual(0, Convert.ToInt32(values[0]));
            Assert.AreEqual(false, values[1]);
        }

        [TestMethod]
        public void SetFromAMemberCallOnACreateObjectResult()
        {
            var values = Execute(
                "Dim keys : keys = CreateObject(\"Scripting.Dictionary\").Keys()\r\n" +
                "collector.Report UBound(keys)\r\n");
            Assert.AreEqual(-1, Convert.ToInt32(values[0]));
        }
    }
}
