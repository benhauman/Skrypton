#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// Inside a Function its name is the return value: Name.Member (args) does not call the function again (issue #161)
        /// </summary>
        [TestMethod] public void ComParity_FunctionNameMemberAccess() => RunParityScript();
    }
}
