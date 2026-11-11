using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport
{
    /// <summary>
    /// NotSupportedException messages say where the problem is: the VBScript line and source and the Skrypton source location that raised it
    /// </summary>
    [TestClass]
    public sealed class NotSupportedExceptionMessageTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object> Values { get; } = new List<object>();
            public void Report(object value) => Values.Add(value);
        }

        [TestMethod]
        public void RuntimeErrorNamesTheValuesAndTheSkryptonSourceLocation()
        {
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", new ValueCollector(), false);
            var exception = Assert.Throws<NotSupportedException>(() => scriptengine.ExecuteStatement("Dim i, d : i = \"xx\" : d = Now()\r\ncollector.Report DateDiff(i, d, d)\r\n"));

            StringAssert.Contains(exception.Message, "DateDiff: unsupported interval 'xx'");
            StringAssert.Matches(exception.Message, new System.Text.RegularExpressions.Regex(@"\[DefaultRuntimeFunctionalityProvider\.cs:\d+ DATEDIFF\]$"));
        }
    }
}
