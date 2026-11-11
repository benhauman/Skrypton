#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/41 - the cached getter invoker must call the call-time object, not the first object of the same type
    /// </summary>
    [TestClass]
    public sealed class Issue41GetterInvokerCacheTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        /// <summary>
        /// An IReflect adapter that answers from its own state ("this"), ignoring the target argument of InvokeMember - like wrappers around host objects do
        /// </summary>
        public abstract class SelfDispatchingBox : IReflect
        {
            protected SelfDispatchingBox(int value)
            {
                V = value;
            }

            public int V { get; }

            protected abstract string Kind { get; }

            public object InvokeMember(string name, BindingFlags invokeAttr, Binder? binder, object? target, object?[]? args, ParameterModifier[]? modifiers, CultureInfo? culture, string[]? namedParameters)
            {
                if (string.Equals(name, "V", StringComparison.OrdinalIgnoreCase))
                    return V;
                if (string.Equals(name, "Kind", StringComparison.OrdinalIgnoreCase))
                    return Kind;
                throw new MissingMemberException(GetType().Name, name);
            }

            public Type UnderlyingSystemType => GetType();
            public FieldInfo? GetField(string name, BindingFlags bindingAttr) => null;
            public FieldInfo[] GetFields(BindingFlags bindingAttr) => [];
            public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) => GetType().GetMember(name, bindingAttr);
            public MemberInfo[] GetMembers(BindingFlags bindingAttr) => GetType().GetMembers(bindingAttr);
            public MethodInfo? GetMethod(string name, BindingFlags bindingAttr) => null;
            public MethodInfo? GetMethod(string name, BindingFlags bindingAttr, Binder? binder, Type[] types, ParameterModifier[]? modifiers) => null;
            public MethodInfo[] GetMethods(BindingFlags bindingAttr) => [];
            public PropertyInfo[] GetProperties(BindingFlags bindingAttr) => GetType().GetProperties(bindingAttr);
            public PropertyInfo? GetProperty(string name, BindingFlags bindingAttr) => GetType().GetProperty(name, bindingAttr);
            public PropertyInfo? GetProperty(string name, BindingFlags bindingAttr, Binder? binder, Type? returnType, Type[] types, ParameterModifier[]? modifiers) => GetType().GetProperty(name, bindingAttr);
        }

        public sealed class BoxA : SelfDispatchingBox
        {
            public BoxA(int value) : base(value) { }
            protected override string Kind => "A";
        }

        public sealed class BoxB : SelfDispatchingBox
        {
            public BoxB(int value) : base(value) { }
            protected override string Kind => "B";
        }

        public sealed class Factory
        {
            public object MakeA(int value) => new BoxA(value);
            public object MakeB(int value) => new BoxB(value);
        }

        private List<object> Execute(string vbs)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.AddObject("ctx", new Factory(), false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        [TestMethod]
        public void TwoInstancesOfTheSameIReflectType()
        {
            var values = Execute(
                "Dim a : Set a = ctx.MakeA(1)\r\n" +
                "Dim b : Set b = ctx.MakeA(2)\r\n" +
                "collector.Report a.V & \",\" & b.V\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual("1,2", values[0]);
        }

        [TestMethod]
        public void InstancesReadInALoop()
        {
            var values = Execute(
                "Dim i, c, s\r\n" +
                "For i = 0 To 3\r\n" +
                "  Set c = ctx.MakeA(i * 10)\r\n" +
                "  s = s & c.V & \";\"\r\n" +
                "Next\r\n" +
                "collector.Report s\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual("0;10;20;30;", values[0]);
        }

        [TestMethod]
        public void TwoIReflectTypesWithTwoInstancesEach()
        {
            var values = Execute(
                "Dim a1 : Set a1 = ctx.MakeA(1)\r\n" +
                "Dim b1 : Set b1 = ctx.MakeB(2)\r\n" +
                "Dim a2 : Set a2 = ctx.MakeA(3)\r\n" +
                "Dim b2 : Set b2 = ctx.MakeB(4)\r\n" +
                "collector.Report a1.Kind & a1.V & b1.Kind & b1.V & a2.Kind & a2.V & b2.Kind & b2.V\r\n");
            Assert.HasCount(1, values);
            Assert.AreEqual("A1B2A3B4", values[0]);
        }
    }
}
