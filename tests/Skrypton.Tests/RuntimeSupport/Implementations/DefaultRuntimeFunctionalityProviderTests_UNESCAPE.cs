
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
			[TestMethod, MyFact]
			public void EmptyResultsInBlankString()
			{
				myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().UNESCAPE(null));
			}

			[TestMethod, MyFact]
			public void NullResultsInNull()
			{
				myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().UNESCAPE(DBNull.Value));
			}

			[TestMethod, MyFact]
			public void PlainString()
			{
				myAssert.AreEqual("test", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().UNESCAPE("test"));
			}

			[TestMethod, MyFact]
			public void ComplexString()
			{
				myAssert.AreEqual("\"Tüst the,th+in%2Bg ć\"", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().UNESCAPE("%22T%FCst%20the%2Cth+in%252Bg%20%u0107%22"));
			}

			[TestMethod, MyFact]
			public void NonEscapedCharacters()
			{
				myAssert.AreEqual("@*_+-./", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().UNESCAPE("@*_+-./"));
			}
		}
	//}
}
