#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// MSXML element dataType / nodeTypedValue: bin.hex and bin.base64 (issue #170)
        /// </summary>
        [TestMethod] public void ComParity_Msxml2BinaryData() => RunParityScript();
    }
}
