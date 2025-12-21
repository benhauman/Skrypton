
using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class ISOBJECT : TestBase
    {
        [TestMethod, MyTheory, MyMemberData("TrueData")]
        public void TrueCases(string description, object value)
        {
            myAssert.True(DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ISOBJECT(value));
        }

        [TestMethod, MyTheory, MyMemberData("FalseData")]
        public void FalseCases(string description, object value)
        {
            myAssert.False(DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().ISOBJECT(value));
        }

        public static IEnumerable<object[]> TrueData
        {
            get
            {
                yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                yield return new object[] { "new Object", new object() };
                yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype() };
            }
        }

        public static IEnumerable<object[]> FalseData
        {
            get
            {
                yield return new object[] { "Empty", null };
                yield return new object[] { "Null", DBNull.Value };
                yield return new object[] { "Zero", 0 };
                yield return new object[] { "Blank string", "" };
                yield return new object[] { "Unintialised array", new object[0] };
            }
        }
    }
    //}
}
