
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
//#using Xunit#;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass] // public static partial class DefaultRuntimeFunctionalityProviderTests
                //{
    public class TYPENAME : TestBase
    {
        //[TestMethod()]
        //public void Cases_Dictionary(string description, object value, string expectedTypeName)
        //{
        //    var entry = Data.First(static x => (string)x[0] == "Scripting Dictionary");
        //    Cases((string)entry[0], entry[1], (string)entry[2]);
        //}
        [TestMethod, MyTheory, MyMemberData(nameof(Data))]
        public void TYPENAMECases(string description, object value, string expectedTypeName)
        {
            Console.WriteLine($"{description} , e:{expectedTypeName}");
            myAssert.AreEqual(expectedTypeName, DefaultRuntimeSupportClassFactoryInstance.Get().TYPENAME(value));
        }

        public static IEnumerable<object[]> Data
        {
            get
            {
                yield return new object[] { "Empty", null, "Empty" };
                yield return new object[] { "Null", DBNull.Value, "Null" };
                yield return new object[] { "Nothing", VBScriptConstants.Nothing, "Nothing" };
                yield return new object[] { "True", true, "Boolean" };
                yield return new object[] { "False", false, "Boolean" };
                yield return new object[] { "Byte", (Byte)1, "Byte" };
                yield return new object[] { "VBScript Integer (Int16)", (Int16)1, "Integer" };
                yield return new object[] { "VBScript Long (Int32)", (Int32)1, "Long" };
                yield return new object[] { "VBScript Double", (double)1, "Double" };
                yield return new object[] { "VBScript Currency (Decimal)", (decimal)1, "Currency" };
                yield return new object[] { "Date", new DateTime(2015, 5, 18, 20, 35, 0), "Date" };
                yield return new object[] { "Date without time component", new DateTime(2015, 5, 18), "Date" };
                yield return new object[] { "VBScript time (ZeroDate with time component)", VBScriptConstants.ZeroDate.Add(new TimeSpan(20, 35, 0)), "Date" };
                yield return new object[] { "Scripting Dictionary", Activator.CreateInstance(typeof(MyScriptingDictionary)), "Dictionary" };
                yield return new object[] { "Translated Class", new exampledefaultpropertytype(), "ExampleDefaultPropertyType" };
                yield return new object[] { "COM Visible Class", new ComVisibleClass(), "ComVisibleClass" };
                yield return new object[] { "Non-COM-Visible Class derived from a COM Visible Class", new NonComVisibleClassDerivedFromComVisibleClass(), "ComVisibleClass" };
                yield return new object[] { "Non-COM-Visible Class", new NonComVisibleClass(), "Object" };
                //lubo:yield return new object[] {
                //lubo:	"WSC Component",
                //lubo:	Interaction.GetObject("script:" + Path.Combine(new DirectoryInfo(".").FullName, @"RuntimeSupport\Implementations\Test.wsc")),
                //lubo:	"Test"
                //lubo:};
                /*lubo:
                -- Load the Windows Script Component file test1.wsc
                -- .wsc files are XML‑based script component definitions that expose COM classes written in VBScript or JScript.
                -- They define <component>, <registration>, <public>, <script> sections, etc.
                lubo: sample WSC file content (test1.wsc) for Interaction.GetObject("script:test1.wsc"):
                    <component>
                     <registration
                       progid="Test1.Component"
                       classid="{12345678-1234-1234-1234-1234567890AB}" />
                     <public>
                       <method name="AddNumbers" />
                     </public>
                     <script language="VBScript">
                       <![CDATA[
                         Function AddNumbers(a, b)
                           AddNumbers = a + b
                         End Function
                       ]]>
                     </script>
                   </component>
                                   */
            }
        }

        [ComVisible(true)]
        private class ComVisibleClass { }

        private class NonComVisibleClassDerivedFromComVisibleClass : ComVisibleClass { }

        private class NonComVisibleClass { }
    }
    //}
}
