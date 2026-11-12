#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// ComParity_ForEach.vbs with the command line tool (issue #112)
        /// </summary>
        [TestMethod]
        public void ComponentParityScriptForEach() => ComponentParityScript("ComParity_ForEach");
    }
}
