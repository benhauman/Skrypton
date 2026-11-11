using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.ScriptControlSupport;

namespace Skrypton.Tests.Issues
{
    /// <summary>
    /// https://github.com/benhauman/Skrypton/issues/47 - IScriptControl.Error and IScriptControl.Procedures
    /// </summary>
    [TestClass]
    public sealed class Issue47ErrorAndProceduresTests : TestBase
    {
        private IScriptControl CreateScriptControl() => CreateScriptControlClass(CreateRuntimeHost(CreateTestHostServices()), CreateScriptControlConfiguration(false, [], []));

        [TestMethod]
        public void ProceduresListsSubsAndFunctions()
        {
            IScriptControl sc = CreateScriptControl();
            sc.AddCode(
                "Sub OnSave()\r\n" +
                "End Sub\r\n" +
                "Function Calc(a, b)\r\n" +
                "  Calc = a + b\r\n" +
                "End Function\r\n" +
                "Sub OnLoad(ByVal x)\r\n" +
                "End Sub\r\n");

            Procedures procedures = sc.Procedures;
            Assert.AreEqual(3, procedures.Count);
            CollectionAssert.AreEqual(new[] { "OnSave", "Calc", "OnLoad" }, procedures.Select(p => p.Name).ToArray());

            Assert.AreEqual(0, procedures["OnSave"].NumArgs);
            Assert.IsFalse(procedures["OnSave"].HasReturnValue);
            Assert.AreEqual(2, procedures["calc"].NumArgs, "lookup is case-insensitive");
            Assert.IsTrue(procedures["CALC"].HasReturnValue);
            Assert.AreEqual("OnLoad", procedures[3].Name, "lookup by 1-based index");
            Assert.ThrowsExactly<KeyNotFoundException>(() => procedures["Missing"]);
        }

        [TestMethod]
        public void ProceduresIsEmptyWithoutCode()
        {
            Assert.AreEqual(0, CreateScriptControl().Procedures.Count);
        }

        [TestMethod]
        public void ErrorDescribesARuntimeError()
        {
            IScriptControl sc = CreateScriptControl();
            sc.AddCode(
                "Sub OnSave()\r\n" +
                "  Dim x\r\n" +
                "  x = 1 / 0\r\n" +
                "End Sub\r\n");
            Assert.AreEqual(0, sc.Error!.Number, "no error before the first call");

            Assert.Throws<Exception>(() => sc.ExecuteStatement("OnSave"));

            Error error = sc.Error!;
            Assert.AreEqual(11, error.Number);
            Assert.AreEqual("Division by zero", error.Description);
            Assert.AreEqual("Microsoft VBScript runtime error", error.Source);
            Assert.AreEqual(3, error.Line, "the VBScript source line");
            Assert.AreEqual("x = 1 / 0", error.Text);

            // A successful call clears the error, as does Clear
            sc.ExecuteStatement("Dim y : y = 1");
            Assert.AreEqual(0, sc.Error!.Number);
            Assert.Throws<Exception>(() => sc.ExecuteStatement("OnSave"));
            Assert.AreEqual(11, sc.Error!.Number);
            sc.Error!.Clear();
            Assert.AreEqual(0, sc.Error!.Number);
            Assert.AreEqual("", sc.Error!.Description);
        }

        [TestMethod]
        public void ErrorDescribesATranslationError()
        {
            IScriptControl sc = CreateScriptControl();
            Assert.Throws<Exception>(() => sc.ExecuteStatement("If Then Else End"));

            Error error = sc.Error!;
            Assert.AreNotEqual(0, error.Number);
            Assert.AreEqual("Microsoft VBScript compilation error", error.Source);
            Assert.AreNotEqual("", error.Description);
        }
    }
}
