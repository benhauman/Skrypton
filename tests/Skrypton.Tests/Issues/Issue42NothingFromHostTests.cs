#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/42 - Nothing returned by a host member must arrive as Nothing (VarType 9), not as Empty (VarType 0)
    /// </summary>
    [TestClass]
    public sealed class Issue42NothingFromHostTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        public interface IRelated
        {
            string Name { get; }
        }

        /// <summary>A plain CLR host whose members return no object</summary>
        public sealed class ClrHost
        {
            [return: MarshalAs(UnmanagedType.IDispatch)]
            public object? GetRelatedObject() => null;

            public IRelated? GetRelatedInterface() => null;
        }

        /// <summary>An IReflect host whose InvokeMember returns VBScriptConstants.Nothing</summary>
        public sealed class ReflectHost : IReflect
        {
            public object? InvokeMember(string name, BindingFlags invokeAttr, Binder? binder, object? target, object?[]? args, ParameterModifier[]? modifiers, CultureInfo? culture, string[]? namedParameters)
            {
                if (string.Equals(name, "GetRelatedObject", StringComparison.OrdinalIgnoreCase))
                    return VBScriptConstants.Nothing;
                throw new MissingMemberException(nameof(ReflectHost), name);
            }

            public Type UnderlyingSystemType => GetType();
            public FieldInfo? GetField(string name, BindingFlags bindingAttr) => null;
            public FieldInfo[] GetFields(BindingFlags bindingAttr) => [];
            public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) => [];
            public MemberInfo[] GetMembers(BindingFlags bindingAttr) => [];
            public MethodInfo? GetMethod(string name, BindingFlags bindingAttr) => null;
            public MethodInfo? GetMethod(string name, BindingFlags bindingAttr, Binder? binder, Type[] types, ParameterModifier[]? modifiers) => null;
            public MethodInfo[] GetMethods(BindingFlags bindingAttr) => [];
            public PropertyInfo[] GetProperties(BindingFlags bindingAttr) => [];
            public PropertyInfo? GetProperty(string name, BindingFlags bindingAttr) => null;
            public PropertyInfo? GetProperty(string name, BindingFlags bindingAttr, Binder? binder, Type? returnType, Type[] types, ParameterModifier[]? modifiers) => null;
        }

        private List<object?> Execute(string vbs)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.AddObject("clr", new ClrHost(), false);
            scriptengine.AddObject("refl", new ReflectHost(), false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        private const string Source_Literal = "Nothing";
        private const string Source_ClrIDispatch = "clr.GetRelatedObject";
        private const string Source_ClrInterface = "clr.GetRelatedInterface";
        private const string Source_IReflect = "refl.GetRelatedObject";

        [TestMethod]
        [DataRow(Source_Literal)]
        [DataRow(Source_ClrIDispatch)]
        [DataRow(Source_ClrInterface)]
        [DataRow(Source_IReflect)]
        public void Direct(string source)
        {
            var values = Execute(
                "Dim r : Set r = Nothing\r\n" +
                "Set r = " + source + "\r\n" +
                "collector.Report VarType(r) & \"|\" & IsObject(r) & \"|\" & (r Is Nothing)\r\n");
            Assert.AreEqual("9|True|True", values[0], source);
        }

        [TestMethod]
        [DataRow(Source_Literal)]
        [DataRow(Source_ClrIDispatch)]
        [DataRow(Source_ClrInterface)]
        [DataRow(Source_IReflect)]
        public void PassedByRefAndByValAndStoredInArray(string source)
        {
            var values = Execute(
                "Function TRef(ByRef o) : TRef = VarType(o) & \"|\" & IsObject(o) & \"|\" & (o Is Nothing) : End Function\r\n" +
                "Function TVal(ByVal o) : TVal = VarType(o) & \"|\" & IsObject(o) & \"|\" & (o Is Nothing) : End Function\r\n" +
                "Dim r : Set r = " + source + "\r\n" +
                "collector.Report TRef(r)\r\n" +
                "collector.Report TVal(r)\r\n" +
                "Dim a(0) : Set a(0) = r\r\n" +
                "collector.Report VarType(a(0)) & \"|\" & IsObject(a(0)) & \"|\" & (a(0) Is Nothing)\r\n");
            Assert.AreEqual("9|True|True", values[0], source + " ByRef");
            Assert.AreEqual("9|True|True", values[1], source + " ByVal");
            Assert.AreEqual("9|True|True", values[2], source + " array");
        }

        [TestMethod]
        public void EmptyIsStillEmpty()
        {
            var values = Execute(
                "Dim e\r\n" +
                "collector.Report VarType(e) & \"|\" & IsObject(e) & \"|\" & IsEmpty(e)\r\n");
            Assert.AreEqual("0|False|True", values[0]);
        }

        [TestMethod]
        [DataRow(Source_Literal)]
        [DataRow(Source_ClrIDispatch)]
        [DataRow(Source_ClrInterface)]
        [DataRow(Source_IReflect)]
        public void MemberAccessRaisesObjectRequired(string source)
        {
            var values = Execute(
                "Dim r : Set r = " + source + "\r\n" +
                "On Error Resume Next\r\n" +
                "Dim x : x = r.Foo\r\n" +
                "collector.Report Err.Number\r\n");
            Assert.AreEqual(424, Convert.ToInt32(values[0]), source);
        }
    }
}
