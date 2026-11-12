#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// CDO.Configuration / CDO.Message (CDOSYS, issue #149) without a mail server: the fields, the bodies and the errors of Send
        /// </summary>
        [TestMethod] public void ComParity_CdoMessage() => RunParityScript();
    }
}
