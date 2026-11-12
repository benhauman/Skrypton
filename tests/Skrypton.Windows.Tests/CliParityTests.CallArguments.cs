#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// ComParity_CallArguments.vbs with the command line tool (issue #120)
        /// </summary>
        [TestMethod]
        public void ComponentParityScriptCallArguments() => ComponentParityScript("ComParity_CallArguments");
    }
}
