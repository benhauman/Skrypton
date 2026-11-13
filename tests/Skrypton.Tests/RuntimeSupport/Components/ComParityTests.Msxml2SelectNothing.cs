#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// MSXML: selectSingleNode of an element and an item out of range are Nothing (issue #179)
        /// </summary>
        [TestMethod] public void ComParity_Msxml2SelectNothing() => RunParityScript();
    }
}
