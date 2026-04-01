
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
	[TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
	//{
		public class UNESCAPE : TestBase
    {
			[TestMethod]
			public void EmptyResultsInBlankString()
			{
				myAssert.AreEqual("", DefaultRuntimeSupportClassFactoryInstance.Get().UNESCAPE(null));
			}

			[TestMethod]
			public void NullResultsInNull()
			{
				myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactoryInstance.Get().UNESCAPE(DBNull.Value));
			}

			[TestMethod]
			public void PlainString()
			{
				myAssert.AreEqual("test", DefaultRuntimeSupportClassFactoryInstance.Get().UNESCAPE("test"));
			}

			[TestMethod]
			public void ComplexString()
			{
				myAssert.AreEqual("\"Tüst the,th+in%2Bg ć\"", DefaultRuntimeSupportClassFactoryInstance.Get().UNESCAPE("%22T%FCst%20the%2Cth+in%252Bg%20%u0107%22"));
			}

			[TestMethod]
			public void NonEscapedCharacters()
			{
				myAssert.AreEqual("@*_+-./", DefaultRuntimeSupportClassFactoryInstance.Get().UNESCAPE("@*_+-./"));
			}
		}
	//}
}
