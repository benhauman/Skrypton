#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// WScript of the demo scripts: Echo writes the text of the value (the demos echo strings)
        /// </summary>
        public sealed class DemoWScript
        {
            private readonly Output _output;

            public DemoWScript(Output output)
            {
                _output = output;
            }

            public void Echo(object? value) => _output.Write(value);
        }

        /// <summary>
        /// The demo scripts of samples\Windows that use only the built-ins (and the file system) - their output was captured with cscript (and
        /// Word, Excel); the demos that use the services of Windows run in Skrypton.Windows.Tests (CliParityTests - issue #142)
        /// </summary>
        [TestMethod]
        [DataRow("Excel-Report")]
        [DataRow("FileSystem-Inventory")]
        [DataRow("Word-Letter")]
        public void DemoScript(string name)
        {
            string script = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.Samples." + name + ".vbs");
            string expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.Samples." + name + ".expected.txt");
            RunParityScript(script, expected, demo: true);
        }
    }
}
