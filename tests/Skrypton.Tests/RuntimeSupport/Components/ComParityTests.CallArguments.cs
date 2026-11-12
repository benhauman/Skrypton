#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// ComParity_CallArguments.vbs (issue #120)
        /// </summary>
        [TestMethod] public void ComParity_CallArguments() => RunParityScript();
    }
}
