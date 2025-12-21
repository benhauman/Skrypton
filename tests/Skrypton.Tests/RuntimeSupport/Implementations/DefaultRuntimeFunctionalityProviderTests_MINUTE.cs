
using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    /// <summary>
    /// This is EXTREMELY close to CDATE.. but not exactly the same (the Null case is the only difference, I think - that CDATE throws an error while this returns Null)
    /// </summary>
    public class MINUTE : TestBase
    {
        [TestMethod, MyTheory, MyMemberData(nameof(SuccessData))]
        public void SuccessCases(string description, object value, object expectedResult)
        {
            myAssert.AreEqual(expectedResult, DefaultRuntimeSupportClassFactoryInstance.Get().MINUTE(value));
        }
        public static IEnumerable<object[]> SuccessData
        {
            get
            {
                yield return new object[] { " 1.Empty", null, 0 };
                yield return new object[] { " 2.Null", DBNull.Value, DBNull.Value };
                yield return new object[] { " 3.Zero", null, 0 };
                yield return new object[] { " 4.Minus one", -1, 0 };
                yield return new object[] { " 5.Minus 400", -400, 0 };
                yield return new object[] { " 6.Minus 400.2", -400.2, 48 };
                yield return new object[] { " 7.Minus 400.8", -400.8, 12 };
                yield return new object[] { " 8.Plus 400.2", 400.2, 48 };
                yield return new object[] { " 9.Plus 400.8", 400.8, 12 };
                yield return new object[] { "10.Plus 40000", 40000, 0 };
                yield return new object[] { "11.String \"-400.2\"", "-400.2", 48 };
                yield return new object[] { "12.String \"2009-10-11\"", "2009-10-11", 0 };
                yield return new object[] { "13.String \"2009-10-11 20:12:44\"", "2009-10-11 20:12:44", 12 };
                yield return new object[] { "14.A Date", new DateTime(2009, 7, 6, 20, 12, 44), 12 };

                yield return new object[] { "Object with default property which is Empty", new exampledefaultpropertytype(), 0 };
                yield return new object[] { "Object with default property which is Null", new exampledefaultpropertytype { result = DBNull.Value }, DBNull.Value };
                yield return new object[] { "Object with default property which is Zero", new exampledefaultpropertytype { result = 0 }, 0 };
                yield return new object[] { "Object with default property which is String \"2009-10-11 20:12:44\"", new exampledefaultpropertytype { result = "2009-10-11 20:12:44" }, 12 };

                // Some bizarre behaviour occurs at the very top end of the supported range - at the very last integer, when 0.9 is present as the time component, then the number
                // of minutes is inconsistent with ANY other value in the acceptable range that has a 0.9 time component; it changes from always being 36 to being 35 at the very
                // last chance.
                yield return new object[] { "Minus 400.9", -400.9, 36 };
                yield return new object[] { "Plus 2000000.9 (approx 2/3 of largest possible positive integer)", 2000000.9, 36 };
                yield return new object[] { "One before the largest positive integer before overflow + 0.9", 2958464.9, 36 };
                yield return new object[] { "Largest positive integer before overflow + 0.9", 2958465.9, 36 }; // lubo old:35 31/12/9999 21:36:00   (en-GB display)
                yield return new object[] { "Most negative possible value with 0.9 time component", -657434.9, 36 };

                // Overflow edge checks
                yield return new object[] { "a.Largest positive integer before overflow", 2958465, 0 };
                yield return new object[] { "b.Largest positive integer before overflow + 0.9", 2958465.9, 36 }; // lubo old:35 31/12/9999 21:36:00   (en-GB display)

                yield return new object[] { "c.Largest positive integer before overflow + 0.99", 2958465.99, 45 };
                yield return new object[] { "d.Largest positive integer before overflow + 0.999", 2958465.999, 58 };
                yield return new object[] { "e.Largest positive integer before overflow + 0.9999", 2958465.9999, 59 };
                yield return new object[] { "f.Largest negative integer before overflow", -657434, 0 };
            }
        }

        [TestMethod, MyTheory, MyMemberData("TypeMismatchData")]
        public void TypeMismatchCases(string description, object value)
        {
            myAssert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().MINUTE(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("ObjectVariableNotSetData")]
        public void ObjectVariableNotSetCases(string description, object value)
        {
            myAssert.Throws<ObjectVariableNotSetException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().MINUTE(value);
            });
        }

        [TestMethod, MyTheory, MyMemberData("OverflowData")]
        public void OverflowCases(string description, object value)
        {
            myAssert.Throws<VBScriptOverflowException>(() =>
            {
                DefaultRuntimeSupportClassFactory.Create(TestCulture).Get().MINUTE(value);
            });
        }



        public static IEnumerable<object[]> TypeMismatchData
        {
            get
            {
                yield return new object[] { "Blank string", "" };
                yield return new object[] { "Object with default property which is a blank string", new exampledefaultpropertytype { result = "" } };
            }
        }

        public static IEnumerable<object[]> ObjectVariableNotSetData
        {
            get
            {
                yield return new object[] { "Nothing", VBScriptConstants.Nothing };
                yield return new object[] { "Object with default property which is Nothing", new exampledefaultpropertytype { result = VBScriptConstants.Nothing } };
            }
        }

        public static IEnumerable<object[]> OverflowData
        {
            get
            {
                yield return new object[] { "Large number (12388888888888.2)", 12388888888888.2 };
                yield return new object[] { "Object with default property which is a large number (12388888888888.2)", new exampledefaultpropertytype { result = 12388888888888.2 } };

                yield return new object[] { "Smallest positive integer that overflows", 2958466 };
                yield return new object[] { "Smallest negative integer that overflows", -657435 };
            }
        }
    }
    //}
}
