using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.CSharpWriter.CodeTranslation;
namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndReDimTranslationTests : TestBase
    {
        [TestMethod]
        public void NonPreserveReDimOfUndeclaredVarInOutermost() // NonPreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope
        {
            string source = @"
                    ReDim a(0)
                ";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void PreserveReDimOfUndeclaredVarInTheOutermost() // PreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope
        {
            string source = @"
                    ReDim Preserve a(0)
                ";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void NonPreserveReDimOfUndeclaredVarInFuncShouldDeclareTheVarInLocal()
        {
            string source = @"
                    Function F1()
                        ReDim a(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void PreserveReDimOfUndeclaredVarInFunctShouldDeclareTheVarInLocal()
        {
            string source = @"
                    Function F1()
                        ReDim Preserve a(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void NonPreserveReDimOfFunctionReturnValue()
        {
            string source = @"
                    Function F1()
                        ReDim F1(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void PreserveReDimOfFunctionReturnValue()
        {
            string source = @"
                    Function F1()
                        ReDim Preserve F1(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
        /// being defined multiple times in the C# code (when the ReDim statements exist within in the outermost scope)
        /// </summary>
        [TestMethod]
        public void RepeatedReDimInOutermostScope1()
        {
            string source = @"
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";

            string actualCs = DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, [], [], [], []);

            myAssert.True(actualCs.Contains("a = null;"), "assign:" + actualCs);
            myAssert.True(actualCs.Contains("internal object a { get; set; }"), "property:" + actualCs);
        }

        /// <summary>
        /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
        /// being defined multiple times in the C# code (when the ReDim statements exist within a function or property)
        /// </summary>
        [TestMethod]
        public void RepeatedReDimInFunction1()
        {
            string source = @"
                    Function F1()
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }
        [TestMethod]
        public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope1()
        {
            string source = @"
                    Dim a
                    ReDim a(0)
                ";
            string[] expected = [
                    "_outer.a = _.NEWARRAY(new object[] { (Int16)0 });"
                ];
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void PreserveReDimOfDeclaredVariableInTheOutermostScope1()
        {
            string source = @"
                    Dim a
                    ReDim Preserve a(0)
                ";
            string[] expected = [
                    "_outer.a = _.RESIZEARRAY(_outer.a, new object[] { (Int16)0 });"
                ];
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void NonPreserveReDimOfDeclaredVariableInFunction1()
        {
            string source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        [TestMethod]
        public void PreserveReDimOfDeclaredVariableInFunction1()
        {
            string source = @"
                    Function F1()
                        Dim a
                        ReDim Preserve a(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
        /// ReDims does not cause any problems (or, in fact, change in behaviour)
        /// </summary>
        [TestMethod]
        public void RepeatedReDimInOutermostScope2()
        {
            string source = @"
                    Dim a
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";

            string text_a_raw = DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, [], [], [], []);

            myAssert.True(text_a_raw.Contains("a = null;"), "assign:" + text_a_raw);
            myAssert.True(text_a_raw.Contains("internal object a { get; set; }"), "property:" + text_a_raw);
        }

        /// <summary>
        /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
        /// ReDims does not cause any problems (or, in fact, change in behaviour)
        /// </summary>
        [TestMethod]
        public void RepeatedReDimInFunction2()
        {
            string source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }

        /// <summary>
        /// A "Dim a()" will result in an explicit array-type variable declaration while a subsequent "ReDim a(0)" will result in an explicit non-array-type
        /// variable declaration (followed by an array initialisation targetting that variable). The non-array-type variable declaration from the ReDim must
        /// be ignored, the array-type declaration from the Dim must take precedence.
        /// </summary>
        [TestMethod]
        public void ReDimFollowingNonDimensionalArrayDimInFunction()
        {
            string source = @"
                    Function F1()
                        Dim a()
                        ReDim a(0)
                    End Function";
            TestCSharpCodeTranslationWithoutScaffolding(source);
        }
        /// <summary>
        /// ReDim will implicitly declare any target variable, if it has not been already declared - this means that a Dim statement that FOLLOWS a ReDim
        /// will result in a "Name redefined" compile time error in VBScript, so all of these cases should result in a translation exception
        /// </summary>
        [TestMethod]
        public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope2()
        {
            string source = @"
                    ReDim a(0)
                    Dim a
                ";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies, [], [], []);
            });
        }

        [TestMethod]
        public void PreserveReDimOfDeclaredVariableInTheOutermostScope2()
        {
            string source = @"
                    ReDim Preserve a(0)
                    Dim a
                ";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies, [], [], []);
            });
        }

        [TestMethod]
        public void NonPreserveReDimOfDeclaredVariableInFunction2()
        {
            string source = @"
                    Function F1()
                        ReDim a(0)
                        Dim a
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies, [], [], []);
            });
        }

        [TestMethod]
        public void PreserveReDimOfDeclaredVariableInFunction2()
        {
            string source = @"
                    Function F1()
                        ReDim Preserve a(0)
                        Dim a
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies, [], [], []);
            });
        }

        /// <summary>
        /// If a ReDim exists for a particular variable before a Dim for the same variable, even if they are not present on a single code branch that may
        /// be executed by a single request, the Dim will still result in a "Name redefined" error being raise
        /// </summary>
        [TestMethod]
        public void ReDimBeforeDimButOnDifferentCodePath()
        {
            string source = @"
                    Function F1()
                        If (True) Then
                            ReDim a(0)
                        Else
                            Dim a
                        End If
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                DefaultCSharpTranslation.GetTranslatedProgramCode(this, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies, [], [], []);
            });
        }

        /// <summary>
        /// While a REDIM statement may be interpreted as explicitly declaring a variable when its target variable has not been declared already in any accessible scope, if there IS
        /// a variable that it might be referencing in a parent scope then the REDIM should NOT be interpreted as explicitly declaring a new variable (even if the variable in the
        /// parent scope was only IMPLICITLY declared - ie. accessed but never DIM'd)
        /// </summary>
        [TestMethod]
        public void ReDimsWithinFuncCanPointToImplicitlyDeclOuterMostScopeVars() // ReDimsWithinFunctionCanPointToImplicitlyDeclaredOuterMostScopeVariables
        {
            string source = @"
                a = 1
                Function F1()
                    ReDim a(2) ' This refers to the implicitly-declared variable ""a"" in the outermost scope
                End Function
                Class C1
                    Private c
                    Function CF1()
                        ReDim a(3) ' This refers to the implicitly-declared variable ""a"" in the outermost scope
                        ReDim b(3) ' There is no reference for this to relate to, so it acts as new explicit variable declaration
                        ReDim c(3) ' This refers to the explicitly-declared variable ""c"" in the containing class
                    End Function
                End Class";
            TestCSharpCodeTranslation(source, ["SKY101"]);
        }
    }
}
