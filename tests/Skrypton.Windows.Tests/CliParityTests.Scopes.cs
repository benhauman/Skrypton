#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// ComParity_Scopes.vbs with the command line tool (issue #130)
        /// </summary>
        [TestMethod]
        public void ComponentParityScriptScopes() => ComponentParityScript("ComParity_Scopes");
    }
}
