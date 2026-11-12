#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// MSXML DOMDocument validateOnParse and resolveExternals (issue #164)
        /// </summary>
        [TestMethod] public void ComParity_Msxml2DocumentProperties() => RunParityScript();
    }
}
