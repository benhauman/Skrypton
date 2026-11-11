#nullable enable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/49 - CLR host objects are bound like their COM dispatch interface (explicit interface members included)
    /// </summary>
    [TestClass]
    public sealed class Issue49InterfaceBindingTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        public interface IThing
        {
            [DispId(0)] string Name { get; }
            [DispId(1)] string Label { get; set; }
            [DispId(2)] [return: MarshalAs(UnmanagedType.IDispatch)] object? Next();
            [DispId(3)] int Add(int a, int b);
            [DispId(4)] string Greet(string name, [Optional, DefaultParameterValue("Hello")] string greeting);
            [DispId(5)] object? NothingOrEmpty();
            [DispId(6)] void Fail();
        }

        /// <summary>Implements its scripting interface explicitly only (so none of the members are public class members)</summary>
        public sealed class ExplicitThing : IThing
        {
            private readonly string _name;
            public ExplicitThing(string name) => _name = name;

            string IThing.Name => _name;
            string IThing.Label { get; set; } = "";
            object? IThing.Next() => null;
            int IThing.Add(int a, int b) => a + b;
            string IThing.Greet(string name, string greeting) => greeting + " " + name;
            object? IThing.NothingOrEmpty() => null;
            void IThing.Fail() => throw new InvalidOperationException("host failure");
        }

        /// <summary>Implements the same interface implicitly</summary>
        public sealed class ImplicitThing : IThing
        {
            public ImplicitThing(string name) => Name = name;
            public string Name { get; }
            public string Label { get; set; } = "";
            public object? Next() => null;
            public int Add(int a, int b) => a + b;
            public string Greet(string name, string greeting = "Hello") => greeting + " " + name;
            public object? NothingOrEmpty() => null;
            public void Fail() => throw new InvalidOperationException("host failure");
        }

        public interface IThingCollection : IEnumerable
        {
            [DispId(1)] int Count { get; }
        }

        public sealed class ExplicitThingCollection : IThingCollection
        {
            private readonly object[] _items;
            public ExplicitThingCollection(params object[] items) => _items = items;
            int IThingCollection.Count => _items.Length;
            IEnumerator IEnumerable.GetEnumerator() => _items.GetEnumerator();
        }

        private List<object?> Execute(string vbs, params (string Name, object Value)[] objects)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            foreach (var (name, value) in objects)
                scriptengine.AddObject(name, value, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        public static IEnumerable<object[]> Things()
        {
            yield return new object[] { "explicit" };
            yield return new object[] { "implicit" };
        }

        private static IThing Create(string kind, string name) => kind == "explicit" ? new ExplicitThing(name) : new ImplicitThing(name);

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void PropertyGetAndSet(string kind)
        {
            IThing thing = Create(kind, "a");
            var values = Execute(
                "collector.Report t.Name\r\n" +
                "t.Label = \"x\"\r\n" +
                "collector.Report t.Label\r\n",
                ("t", thing));
            CollectionAssert.AreEqual(new object[] { "a", "x" }, values.ToArray(), kind);
            Assert.AreEqual("x", thing.Label);
        }

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void MethodsWithArgumentsAndOptionalArguments(string kind)
        {
            var values = Execute(
                "collector.Report t.Add(2, 3)\r\n" +
                "collector.Report t.Greet(\"Bob\")\r\n" +
                "collector.Report t.Greet(\"Bob\", \"Hi\")\r\n",
                ("t", Create(kind, "a")));
            Assert.AreEqual(5, Convert.ToInt32(values[0]), kind);
            Assert.AreEqual("Hello Bob", values[1], kind);
            Assert.AreEqual("Hi Bob", values[2], kind);
        }

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void ArgumentsAreCoerced(string kind)
        {
            var values = Execute("collector.Report t.Add(\"5\", 1)\r\n", ("t", Create(kind, "a")));
            Assert.AreEqual(6, Convert.ToInt32(values[0]), kind);
        }

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void DispIdZeroIsTheDefaultMember(string kind)
        {
            var values = Execute("collector.Report \"\" & t\r\n", ("t", Create(kind, "a")));
            Assert.AreEqual("a", values[0], kind);
        }

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void NullReturnsAreNothingOrEmpty(string kind)
        {
            var values = Execute(
                "Dim n : Set n = t.Next\r\n" +
                "collector.Report VarType(n) & \"|\" & (n Is Nothing)\r\n" +
                "Dim e : e = t.NothingOrEmpty\r\n" +
                "collector.Report VarType(e)\r\n",
                ("t", Create(kind, "a")));
            Assert.AreEqual("9|True", values[0], kind + ": MarshalAs(IDispatch) null is Nothing");
            Assert.AreEqual(0, Convert.ToInt32(values[1]), kind + ": a plain object null is Empty");
        }

        [TestMethod]
        [DynamicData(nameof(Things))]
        public void HostExceptionKeepsItsMessage(string kind)
        {
            var exception = Assert.Throws<Exception>(() => Execute("t.Fail\r\n", ("t", Create(kind, "a"))));
            StringAssert.Contains(exception.ToString(), "host failure");
        }

        [TestMethod]
        public void TwoInstancesOfTheSameType()
        {
            var values = Execute(
                "collector.Report a.Name & b.Name\r\n",
                ("a", new ExplicitThing("1")), ("b", new ExplicitThing("2")));
            Assert.AreEqual("12", values[0]);
        }

        [TestMethod]
        public void ForEachAndCountOnAnExplicitHostCollection()
        {
            var values = Execute(
                "Dim item, s\r\n" +
                "For Each item In c\r\n" +
                "  s = s & item & \";\"\r\n" +
                "Next\r\n" +
                "collector.Report c.Count & \":\" & s\r\n",
                ("c", new ExplicitThingCollection("x", "y")));
            Assert.AreEqual("2:x;y;", values[0]);
        }
    }
}
