#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// Hex and octal literals: the values and types of VBScript (issue #169)
        /// </summary>
        [TestMethod] public void ComParity_HexLiterals() => RunParityScript();
    }
}
