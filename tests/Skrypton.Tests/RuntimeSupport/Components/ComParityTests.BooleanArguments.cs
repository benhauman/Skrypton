#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// A number or a string passed for a Boolean (or numeric) parameter of a built-in is converted as COM converts it: 0, 1, -1, "True", "False"
        /// (issue #138) - a value that can not be converted is a Type mismatch
        /// </summary>
        [TestMethod] public void ComParity_BooleanArguments() => RunParityScript();
    }
}
