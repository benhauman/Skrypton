using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/40 - VBScript And / Or never short-circuit, both operands are always evaluated
    /// </summary>
    [TestClass]
    public sealed class Issue40AndOrNoShortCircuitTests : TestBase
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
        public void BothOperandsAreEvaluatedInIfConditions()
        {
            var values = Execute(
                "Dim n : n = 0\r\n" +
                "Function F() : n = n + 1 : F = True : End Function\r\n" +
                "If False And F() Then\r\n" +
                "End If\r\n" +
                "If True Or F() Then\r\n" +
                "End If\r\n" +
                "collector.Report n\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual(2, Convert.ToInt32(values[0]));
        }

        [TestMethod]
        public void BothOperandsAreEvaluatedOutsideIfConditions()
        {
            var values = Execute(
                "Dim n : n = 0\r\n" +
                "Function F() : n = n + 1 : F = True : End Function\r\n" +
                "Dim x\r\n" +
                "x = False And F()\r\n" +
                "x = True Or F()\r\n" +
                "collector.Report n\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual(2, Convert.ToInt32(values[0]));
        }

        [TestMethod]
        [DataRow("False And o.GetValue(\"X\") = \"Y\"")]
        [DataRow("True Or o.GetValue(\"X\") = \"Y\"")]
        public void ErrorInRightOperandIsRaised(string condition)
        {
            string vbs =
                "Dim o : Set o = Nothing\r\n" +
                "On Error Resume Next\r\n" +
                "If " + condition + " Then\r\n" +
                "End If\r\n" +
                "collector.Report Err.Number\r\n";
            var values = Execute(vbs);
            Assert.HasCount(1, values);
            Assert.AreEqual(424, Convert.ToInt32(values[0]), "Object required");
        }

        [TestMethod]
        [DataRow("6 And 3", 2)]
        [DataRow("6 Or 3", 7)]
        [DataRow("True And 6", 6)]
        [DataRow("False Or 5", 5)]
        public void AndOrAreBitwise(string expression, int expected)
        {
            var values = Execute("collector.Report " + expression + "\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual(expected, Convert.ToInt32(values[0]), expression);
        }

        [TestMethod]
        [DataRow("Null And False", "Boolean")]
        [DataRow("False And Null", "Boolean")]
        [DataRow("Null And True", "Null")]
        [DataRow("Null And Null", "Null")]
        public void NullHandling(string expression, string expectedTypeName)
        {
            var values = Execute("collector.Report TypeName(" + expression + ")\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual(expectedTypeName, values[0], expression);
        }
    }
}
