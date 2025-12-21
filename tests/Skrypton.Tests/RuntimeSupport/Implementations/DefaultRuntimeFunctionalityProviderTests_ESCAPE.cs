
using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
	[TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
	//{
		public class ESCAPE : TestBase
    {
			[TestMethod, MyFact]
			public void EmptyResultsInBlankString()
			{
				myAssert.AreEqual("", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ESCAPE(null));
			}

			[TestMethod, MyFact]
			public void NullResultsInNull()
			{
				myAssert.AreEqual(DBNull.Value, DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ESCAPE(DBNull.Value));
			}

			[TestMethod, MyFact]
			public void PlainString()
			{
				myAssert.AreEqual("test", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ESCAPE("test"));
			}

			[TestMethod, MyFact]
			public void ComplexString()
			{
				myAssert.AreEqual("%22T%FCst%20the%2Cth+in%252Bg%20%u0107%22", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ESCAPE("\"Tüst the,th+in%2Bg ć\""));
			}

			[TestMethod, MyFact]
			public void NonEscapedCharacters()
			{
				myAssert.AreEqual("@*_+-./", DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ESCAPE("@*_+-./"));
			}
		}
	//}
}
