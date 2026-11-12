#nullable enable
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// A Dim / Const in a Sub / Function with the name of a global declared further down (issue #130)
        /// </summary>
        [TestMethod] public void ComParity_Scopes() => RunParityScript();

        /// <summary>
        /// A name declared twice in the same scope is still "Name redefined" (issue #130)
        /// </summary>
        [TestMethod]
        [DataRow("Sub S()\r\n    Dim x\r\n    Dim x\r\nEnd Sub\r\n")]
        [DataRow("Dim x\r\nSub S()\r\nEnd Sub\r\nDim x\r\n")]
        public void NameDeclaredTwiceInTheSameScopeIsRedefined(string script)
        {
            IScriptControl scriptengine = CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), VBScriptRuleConfiguration());
            var exception = Assert.Throws<Exception>(() => scriptengine.ExecuteStatement(script));
            StringAssert.Contains(exception.ToString(), "Name redefined");
        }
    }
}
