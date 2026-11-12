#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// ComParity_Dates.vbs (issue #124)
        /// </summary>
        [TestMethod] public void ComParity_Dates() => RunParityScript();
    }
}
