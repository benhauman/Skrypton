#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// Omitted arguments (issue #154): the first argument left out of a call without brackets, TypeName / VarType Error, the defaults of a built-in
        /// </summary>
        [TestMethod] public void ComParity_OmittedArguments() => RunParityScript();
    }
}
