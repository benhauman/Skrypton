using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/66 - date literals must compile and run when a script is executed
    /// </summary>
    [TestClass]
    public sealed class Issue66DateLiteralTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        private object Execute(string vbs)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement(vbs);
            Assert.HasCount(1, collector.Values);
            return collector.Values[0];
        }

        [TestMethod]
        public void DateLiteralsAsFunctionArguments()
        {
            Assert.AreEqual(1, Convert.ToInt32(Execute("collector.Report DateDiff(\"d\", #2020-01-01#, #2020-01-02#)\r\n")));
            Assert.AreEqual(2015, Convert.ToInt32(Execute("collector.Report Year(#29 May 2015#)\r\n")));
        }

        [TestMethod]
        public void DateLiteralAsAValue()
        {
            Assert.AreEqual(new DateTime(2020, 1, 1), Execute("Dim d : d = #2020-01-01#\r\ncollector.Report d\r\n"));
        }

        [TestMethod]
        public void DateLiteralInAComparison()
        {
            Assert.AreEqual(true, Execute("Dim d : d = DateSerial(2015, 5, 29)\r\ncollector.Report (d = #29 May 2015#)\r\n"));
        }
    }
}
