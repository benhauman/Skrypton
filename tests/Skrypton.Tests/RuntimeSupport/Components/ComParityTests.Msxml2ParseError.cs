#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// DOMDocument.parseError after loadXML (issue #173)
        /// </summary>
        [TestMethod] public void ComParity_Msxml2ParseError() => RunParityScript();
    }
}
