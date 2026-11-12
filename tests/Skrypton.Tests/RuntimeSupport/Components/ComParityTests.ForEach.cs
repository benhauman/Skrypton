#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// ComParity_ForEach.vbs (issue #112)
        /// </summary>
        [TestMethod] public void ComParity_ForEach() => RunParityScript();
    }
}
