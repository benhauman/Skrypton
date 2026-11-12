#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// FileSystemObject text files: ANSI, UTF-16 and the default format, reading and writing (issue #163)
        /// </summary>
        [TestMethod] public void ComParity_FileSystemObjectEncodings() => RunParityScript();
    }
}
