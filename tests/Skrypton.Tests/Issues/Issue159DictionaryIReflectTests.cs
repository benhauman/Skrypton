#nullable enable
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/159 - the built-in Scripting.Dictionary is an IReflect (ScriptonIReflectOnClrType) as the
    /// other built-ins: a client reaches its members (with the DISPIDs of scrrun's IDictionary) through IReflect
    /// </summary>
    [TestClass]
    public sealed class Issue159DictionaryIReflectTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        [TestMethod]
        public void TheDictionaryIsAnIReflectWithTheDispIdsOfScrrun()
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            scriptengine.ExecuteStatement("Dim d : Set d = CreateObject(\"Scripting.Dictionary\")\r\nd.Add \"a\", 1\r\ncollector.Report d\r\n");

            object? dictionary = collector.Values.Single();
            if (dictionary is Skrypton.RuntimeSupport.ScriptDispatchWrapper wrapper)
                dictionary = wrapper.WrappedObject;
            var reflect = dictionary as IReflect;
            Assert.IsNotNull(reflect, "the dictionary is an IReflect");

            // the DISPIDs of scrrun's IDictionary (IDispatch.GetIDsOfNames of the COM Scripting.Dictionary)
            var expected = new Dictionary<string, int>
            {
                ["Item"] = 0, ["Add"] = 1, ["Count"] = 2, ["Exists"] = 3, ["Items"] = 4, ["Key"] = 5, ["Keys"] = 6, ["Remove"] = 7, ["RemoveAll"] = 8, ["CompareMode"] = 9, ["_NewEnum"] = -4,
            };
            foreach (KeyValuePair<string, int> member in expected)
            {
                MemberInfo[] members = reflect.GetMember(member.Key, BindingFlags.Public | BindingFlags.Instance);
                Assert.IsNotEmpty(members, member.Key);
                Assert.AreEqual(member.Value, members[0].GetCustomAttribute<DispIdAttribute>()!.Value, member.Key);
            }

            Assert.AreEqual(1, reflect.InvokeMember("Count", BindingFlags.GetProperty, null, dictionary!, null, null, null, null));
            Assert.AreEqual(true, reflect.InvokeMember("Exists", BindingFlags.InvokeMethod, null, dictionary!, new object[] { "a" }, null, null, null));
        }
    }
}
