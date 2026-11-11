using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/50 - built-in Scripting.Dictionary with scrrun semantics (no host factory, no COM)
    /// </summary>
    [TestClass]
    public sealed class Issue50ScriptingDictionaryTests : TestBase
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
            scriptengine.ExecuteStatement("Dim d : Set d = CreateObject(\"Scripting.Dictionary\")\r\n" + vbs);
            return collector.Values;
        }

        private object Single(string vbs)
        {
            List<object> values = Execute(vbs);
            Assert.HasCount(1, values);
            return values[0];
        }

        public sealed class HostDictionary
        {
            public int Count => 42;
        }

        [TestMethod]
        public void AHostRegistrationOverridesTheBuiltInDictionary()
        {
            var collector = new ValueCollector();
            ScriptControlClass scriptControlClass = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptControlClass.TestSetDefaultRuntimeFunctionalityProviderSetup(provider => provider.RegisterObjectCreateFactory("Scripting.Dictionary", _ => new HostDictionary()));
            IScriptControl scriptengine = scriptControlClass;
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("Dim d : Set d = CreateObject(\"Scripting.Dictionary\")\r\ncollector.Report d.Count\r\n");
            Assert.AreEqual(42, Convert.ToInt32(collector.Values[0]));
        }

        [TestMethod]
        public void TypeNameCountAndRemoveAll()
        {
            Assert.AreEqual("Dictionary|2|0", Single(
                "d.Add \"a\", 1 : d.Add \"b\", 2\r\n" +
                "Dim c : c = d.Count\r\n" +
                "d.RemoveAll\r\n" +
                "collector.Report TypeName(d) & \"|\" & c & \"|\" & d.Count\r\n"));
        }

        [TestMethod]
        public void AddingADuplicateKeyRaises457()
        {
            Assert.AreEqual(457, Convert.ToInt32(Single(
                "On Error Resume Next\r\n" +
                "d.Add \"a\", 1 : d.Add \"a\", 2\r\n" +
                "collector.Report Err.Number\r\n")));
        }

        [TestMethod]
        public void ReadingAMissingKeyAddsItWithEmptyAndAssignmentAddsOrReplaces()
        {
            Assert.AreEqual("True|True|1|x|y|2", Single(
                "Dim v : v = d(\"missing\")\r\n" +
                "Dim r : r = IsEmpty(v) & \"|\" & d.Exists(\"missing\") & \"|\" & d.Count\r\n" +
                "d(\"k\") = \"x\" : r = r & \"|\" & d(\"k\")\r\n" +
                "d(\"k\") = \"y\" : r = r & \"|\" & d.Item(\"k\") & \"|\" & d.Count\r\n" +
                "collector.Report r\r\n"));
        }

        [TestMethod]
        public void KeysItemsAndForEachAreInInsertionOrder()
        {
            Assert.AreEqual("c,a,b|3,1,2|c;a;b;", Single(
                "d.Add \"c\", 3 : d.Add \"a\", 1 : d.Add \"b\", 2\r\n" +
                "Dim k, s\r\n" +
                "For Each k In d : s = s & k & \";\" : Next\r\n" +
                "collector.Report Join(d.Keys(), \",\") & \"|\" & Join(d.Items(), \",\") & \"|\" & s\r\n"));
        }

        [TestMethod]
        public void RenamingAKeyKeepsItsPosition()
        {
            Assert.AreEqual("a,x,c|2", Single(
                "d.Add \"a\", 1 : d.Add \"b\", 2 : d.Add \"c\", 3\r\n" +
                "d.Key(\"b\") = \"x\"\r\n" +
                "collector.Report Join(d.Keys(), \",\") & \"|\" & d(\"x\")\r\n"));
        }

        [TestMethod]
        [DataRow("d.Key(\"missing\") = \"x\"")]
        [DataRow("d.Remove \"missing\"")]
        public void MissingKeyForKeyOrRemoveRaises32811(string statement)
        {
            Assert.AreEqual(32811, Convert.ToInt32(Single(
                "On Error Resume Next\r\n" +
                statement + "\r\n" +
                "collector.Report Err.Number\r\n")));
        }

        [TestMethod]
        public void CompareMode()
        {
            Assert.AreEqual("False|True|5", Single(
                "Dim r : d.Add \"A\", 1 : r = d.Exists(\"a\")\r\n" +
                "Set d = CreateObject(\"Scripting.Dictionary\")\r\n" +
                "d.CompareMode = 1 : d.Add \"A\", 1 : r = r & \"|\" & d.Exists(\"a\")\r\n" +
                "On Error Resume Next\r\n" +
                "d.CompareMode = 0\r\n" +
                "collector.Report r & \"|\" & Err.Number\r\n"));
        }

        [TestMethod]
        public void NumericAndStringKeysAreDistinctAndEmptyAndObjectKeysAreAllowed()
        {
            Assert.AreEqual("4|x|y|e|o", Single(
                "Dim o : Set o = CreateObject(\"Scripting.Dictionary\")\r\n" +
                "Dim emptyKey\r\n" +
                "d.Add 1, \"x\" : d.Add \"1\", \"y\" : d.Add emptyKey, \"e\" : d.Add o, \"o\"\r\n" +
                "collector.Report d.Count & \"|\" & d(1) & \"|\" & d(\"1\") & \"|\" & d(emptyKey) & \"|\" & d(o)\r\n"));
        }
    }
}
