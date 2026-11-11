using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/39 - operators of the same precedence level must be left-associative
    /// </summary>
    [TestClass]
    public sealed class Issue39OperatorPrecedenceTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        [TestMethod]
        [DataRow("1 - 2 + 3 - 4", "_.SUBT(_.ADD(_.SUBT((Int16)1, (Int16)2), (Int16)3), (Int16)4)")]
        [DataRow("5 - 1 + 1", "_.ADD(_.SUBT((Int16)5, (Int16)1), (Int16)1)")]
        [DataRow(@"7 \ 2 * 2", "_.INTDIV((Int16)7, _.MULT((Int16)2, (Int16)2))")]
        [DataRow(@"2 * 3 \ 4", "_.INTDIV(_.MULT((Int16)2, (Int16)3), (Int16)4)")]
        [DataRow("8 * 2 / 4", "_.DIV(_.MULT((Int16)8, (Int16)2), (Int16)4)")]
        public void Grouping(string vbsExpression, string expectedCs)
        {
            string cs = DefaultTranslator.TranslateWithoutScaffolding(TestCulture, "x = " + vbsExpression + "\r\n", new NonNullImmutableList<string>(), new List<ExternalMemberMethodInfo>(), new[] { "SKY101" });
            Console.WriteLine(cs);
            StringAssert.Contains(cs, expectedCs);
        }

        [TestMethod]
        [DataRow("1 - 2 + 3 - 4", -2.0)]
        [DataRow("5 - 1 + 1", 5.0)]
        [DataRow("10 - 2 - 3", 5.0)]
        [DataRow(@"7 \ 2 * 2", 1.0)]
        [DataRow(@"2 * 3 \ 4", 1.0)]
        [DataRow("8 * 2 / 4", 4.0)]
        [DataRow("8 / 2 * 2", 8.0)]
        [DataRow("2 ^ 3 ^ 2", 64.0)]
        [DataRow("9 Mod 4 * 2", 1.0)]
        [DataRow("1 + 2 * 3 - 4 / 2", 5.0)]
        [DataRow(@"-7 \ 2", -3.0)]
        [DataRow(@"2.5 \ 1", 2.0)]
        [DataRow(@"3.5 \ 1", 4.0)]
        [DataRow("2 ^ -1", 0.5)]
        [DataRow("UBound(Split(\"\", \",\")) - LBound(Split(\"\", \",\")) + 1", 0.0)]
        public void Value(string vbsExpression, double expected)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("collector.Report " + vbsExpression + "\r\n");

            Assert.HasCount(1, collector.Values);
            Assert.AreEqual(expected, Convert.ToDouble(collector.Values[0]), vbsExpression);
        }
    }
}
