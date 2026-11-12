#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Windows.Tests
{
    public sealed partial class CliParityTests
    {
        /// <summary>
        /// ComParity_Dates.vbs with the command line tool (issue #124)
        /// </summary>
        [TestMethod]
        public void ComponentParityScriptDates() => ComponentParityScript("ComParity_Dates");
    }
}
