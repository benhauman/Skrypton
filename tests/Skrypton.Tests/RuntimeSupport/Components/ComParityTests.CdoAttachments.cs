#nullable enable
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public sealed partial class ComParityTests
    {
        /// <summary>
        /// CDO.Message.AddAttachment and Attachments (issue #175)
        /// </summary>
        [TestMethod] public void ComParity_CdoAttachments() => RunParityScript();
    }
}
