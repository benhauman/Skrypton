
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
	[TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
	//{
		public class RANDOMIZE_and_RND : TestBase
    {
			[TestMethod, MyFact]
			public void RandomizeSeedReturnsConsistentValuesFirstTimeItIsUsedForRuntimeSupportFactoryInstance()
			{
				const int seed = 123;
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE(seed);
					value1 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE(seed);
					value2 = _.RND();
				}
				myAssert.AreEqual(value1, value2);
			}

			[TestMethod, MyFact]
			public void RandomizeSeedReturnsDifferentSequencesIfUsedWithinSameRuntimeSupportFactoryInstance()
			{
				const int seed = 123;
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE(seed);
					value1 = _.RND();

					_.RANDOMIZE(seed);
					value2 = _.RND();
				}
				myAssert.NotEqual(value1, value2);
			}

			[TestMethod, MyFact]
			public void CallingRndWithZeroReturnsPreviousNumber()
			{
				float value1, value2;
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					value1 = _.RND();
					value2 = _.RND(0);
				}
				myAssert.AreEqual(value1, value2);
			}

			[TestMethod, MyFact]
			public void CallingRndWithNegativeValueSetsSeedToThatValue()
			{
				const int seed = -123;
				float[] values1, values2;
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					values1 = new[] { _.RND(seed), _.RND(), _.RND() };
					values2 = new[] { _.RND(seed), _.RND(), _.RND() };
				}
				myAssert.AreEqual(values1, values2);
			}

			/// <summary>
			/// The precision of the RANDOMIZE seed is limited to a Single (in VBScript parlance, which I think is equivalent to .NET) - so the extra digit on 1.1111111 does not make any
			/// difference compared to 1.111111 (though going one smaller at 1.11111 WILL result in a different sequence being generated)
			/// </summary>
			[TestMethod, MyFact]
			public void RandomizeSeedValueHasLimitedPrecision()
			{
				// The values
				//  1.111111
				// and
				//  1.1111111
				// will result in the same random number streams being generated
				float value1, value2, value3;
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE("1.111111");
					value1 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE("1.1111111");
					value2 = _.RND();
				}
				using (var _ = DefaultRuntimeSupportClassFactory.Create(TestCulture).Get())
				{
					_.RANDOMIZE("1.11111");
					value3 = _.RND();
				}
				myAssert.AreEqual(value1, value2);
            myAssert.NotEqual(value1, value3);
			}
		}
	//}
}
