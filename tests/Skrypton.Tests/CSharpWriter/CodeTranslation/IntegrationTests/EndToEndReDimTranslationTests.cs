using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndReDimTranslationTests : TestBase
    {
        //public class UndeclaredVariables
        //{
        [TestMethod, MyFact]
        public void NonPreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope()
        {
            var source = @"
                    ReDim a(0)
                ";
            var expected = new[] {
                    "_outer.a = _.NEWARRAY(new object[] { (Int16)0 });"
                };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfUndeclaredVariableInTheOutermostScopeShouldImplicitlyDeclareTheVariableInOutermostScope()
        {
            var source = @"
                    ReDim Preserve a(0)
                ";
            var expected = new[] {
                    "_outer.a = _.RESIZEARRAY(_outer.a, new object[] { (Int16)0 });"
                };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void NonPreserveReDimOfUndeclaredVariableInFunctionShouldImplicitlyDeclareTheVariableInLocalScope()
        {
            var source = @"
                    Function F1()
                        ReDim a(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfUndeclaredVariableInFunctionShouldImplicitlyDeclareTheVariableInLocalScope()
        {
            var source = @"
                    Function F1()
                        ReDim Preserve a(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.RESIZEARRAY(a, new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void NonPreserveReDimOfFunctionReturnValue()
        {
            var source = @"
                    Function F1()
                        ReDim F1(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      retVal1 = _.NEWARRAY(new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfFunctionReturnValue()
        {
            var source = @"
                    Function F1()
                        ReDim Preserve F1(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      retVal1 = _.RESIZEARRAY(retVal1, new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
        /// being defined multiple times in the C# code (when the ReDim statements exist within in the outermost scope)
        /// </summary>
        [TestMethod, MyFact]
        public void RepeatedReDimInOutermostScope1()
        {
            var source = @"
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";

            var trimmedTranslatedStatements = DefaultTranslator.Translate(TestCulture, source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable)
                .Select(s => s.Content.Trim())
                .ToArray();

            myAssert.AreEqual(1, trimmedTranslatedStatements.Count(s => s == "a = null;"));
            myAssert.AreEqual(1, trimmedTranslatedStatements.Count(s => s == "public object a { get; set; }"));
        }

        /// <summary>
        /// This test is just to ensure that multiple ReDim statements for the same otherwise-undeclared variable do not result in that variable
        /// being defined multiple times in the C# code (when the ReDim statements exist within a function or property)
        /// </summary>
        [TestMethod, MyFact]
        public void RepeatedReDimInFunction1()
        {
            var source = @"
                    Function F1()
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(new object[] { (Int16)0 });
                      a = _.NEWARRAY(new object[] { (Int16)1 });
                      a = _.NEWARRAY(new object[] { (Int16)2 });
                      return retVal1;
                   }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
        //}
        //[TestClass]
        //public class DeclaredVariables
        //{
        [TestMethod, MyFact]
        public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope1()
        {
            var source = @"
                    Dim a
                    ReDim a(0)
                ";
            var expected = new[] {
                    "_outer.a = _.NEWARRAY(new object[] { (Int16)0 });"
                };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfDeclaredVariableInTheOutermostScope1()
        {
            var source = @"
                    Dim a
                    ReDim Preserve a(0)
                ";
            var expected = new[] {
                    "_outer.a = _.RESIZEARRAY(_outer.a, new object[] { (Int16)0 });"
                };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void NonPreserveReDimOfDeclaredVariableInFunction1()
        {
            var source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfDeclaredVariableInFunction1()
        {
            var source = @"
                    Function F1()
                        Dim a
                        ReDim Preserve a(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.RESIZEARRAY(a, new object[] { (Int16)0 });
                      return retVal1;
                    }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
        /// ReDims does not cause any problems (or, in fact, change in behaviour)
        /// </summary>
        [TestMethod, MyFact]
        public void RepeatedReDimInOutermostScope2()
        {
            var source = @"
                    Dim a
                    ReDim a(0)
                    ReDim a(1)
                    ReDim a(2)";

            var trimmedTranslatedStatements = DefaultTranslator.Translate(TestCulture, source, new string[0], OuterScopeBlockTranslator.OutputTypeOptions.Executable)
                .Select(s => s.Content.Trim())
                .ToArray();

            myAssert.AreEqual(1, trimmedTranslatedStatements.Count(s => s == "a = null;"));
            myAssert.AreEqual(1, trimmedTranslatedStatements.Count(s => s == "public object a { get; set; }"));
        }

        /// <summary>
        /// This is almost identical to the corresponding test in the UndeclaredVariables class but it ensure that a Dim statement before the repeated
        /// ReDims does not cause any problems (or, in fact, change in behaviour)
        /// </summary>
        [TestMethod, MyFact]
        public void RepeatedReDimInFunction2()
        {
            var source = @"
                    Function F1()
                        Dim a
                        ReDim a(0)
                        ReDim a(1)
                        ReDim a(2)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = null;
                      a = _.NEWARRAY(new object[] { (Int16)0 });
                      a = _.NEWARRAY(new object[] { (Int16)1 });
                      a = _.NEWARRAY(new object[] { (Int16)2 });
                      return retVal1;
                   }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A "Dim a()" will result in an explicit array-type variable declaration while a subsequent "ReDim a(0)" will result in an explicit non-array-type
        /// variable declaration (followed by an array initialisation targetting that variable). The non-array-type variable declaration from the ReDim must
        /// be ignored, the array-type declaration from the Dim must take precedence.
        /// </summary>
        [TestMethod, MyFact]
        public void ReDimFollowingNonDimensionalArrayDimInFunction()
        {
            var source = @"
                    Function F1()
                        Dim a()
                        ReDim a(0)
                    End Function";
            var expected = @"
                    public object f1()
                    {
                      object retVal1 = null;
                      object a = (object[])null;
                      a = _.NEWARRAY(new object[] { (Int16)0 });
                      return retVal1;
                   }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
        //}
        //[TestClass]
        /// <summary>
        /// ReDim will implicitly declare any target variable, if it has not been already declared - this means that a Dim statement that FOLLOWS a ReDim
        /// will result in a "Name redefined" compile time error in VBScript, so all of these cases should result in a translation exception
        /// </summary>
        //public class PrecedingExplicitVariableDeclarations
        //{
        [TestMethod, MyFact]
        public void NonPreserveReDimOfDeclaredVariableInTheOutermostScope2()
        {
            var source = @"
                    ReDim a(0)
                    Dim a
                ";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfDeclaredVariableInTheOutermostScope2()
        {
            var source = @"
                    ReDim Preserve a(0)
                    Dim a
                ";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        [TestMethod, MyFact]
        public void NonPreserveReDimOfDeclaredVariableInFunction2()
        {
            var source = @"
                    Function F1()
                        ReDim a(0)
                        Dim a
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        [TestMethod, MyFact]
        public void PreserveReDimOfDeclaredVariableInFunction2()
        {
            var source = @"
                    Function F1()
                        ReDim Preserve a(0)
                        Dim a
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }

        /// <summary>
        /// If a ReDim exists for a particular variable before a Dim for the same variable, even if they are not present on a single code branch that may
        /// be executed by a single request, the Dim will still result in a "Name redefined" error being raise
        /// </summary>
        [TestMethod, MyFact]
        public void ReDimBeforeDimButOnDifferentCodePath()
        {
            var source = @"
                    Function F1()
                        If (True) Then
                            ReDim a(0)
                        Else
                            Dim a
                        End If
                    End Function";
            myAssert.Throws<NameRedefinedException>(() =>
            {
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
            });
        }
        ///}
        ///
        ///[TestClass]
        ///public class EndToEndReDimTranslationTests
        ///{
        /// <summary>
        /// While a REDIM statement may be interpreted as explicitly declaring a variable when its target variable has not been declared already in any accessible scope, if there IS
        /// a variable that it might be referencing in a parent scope then the REDIM should NOT be interpreted as explicitly declaring a new variable (even if the variable in the
        /// parent scope was only IMPLICITLY declared - ie. accessed but never DIM'd)
        /// </summary>
        [TestMethod, MyFact]
        public void ReDimsWithinFunctionCanPointToImplicitlyDeclaredOuterMostScopeVariables()
        {
            var source = @"
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
            var expected = @"
                _env.a = (Int16)1;
                public object f1()
                {
                    object retVal1 = null;
                    _env.a = _.NEWARRAY(new object[] { (Int16)2 }); // This refers to the implicitly-declared variable ""a"" in the outermost scope
                    return retVal1;
                }
                [ComVisible(true)]
                [SourceClassName(""C1"")]
                public sealed class c1
                {
                    private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
                    private readonly EnvironmentReferences _env;
                    private readonly GlobalReferences _outer;
                    public c1(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences outer)
                    {
						_ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
						_env = env ?? throw new ArgumentNullException(nameof(env));
						_outer = outer ?? throw new ArgumentNullException(nameof(outer));
                        c = null;
                    }
                    private object c { get; set; }
                    public object cf1()
                    {
                        object retVal2 = null;
                        object b = null;
                        _env.a = _.NEWARRAY(new object[] { (Int16)3 }); // This refers to the implicitly-declared variable ""a"" in the outermost scope
                        b = _.NEWARRAY(new object[] { (Int16)3 }); // There is no reference for this to relate to, so it acts as new explicit variable declaration
                        c = _.NEWARRAY(new object[] { (Int16)3 }); // This refers to the explicitly-declared variable ""c"" in the containing class
                        return retVal2;
                    }
                }";
            myAssert.AreEqual(
                SplitOnNewLinesSkipFirstLineAndTrimAll(expected).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        private static IEnumerable<string> SplitOnNewLinesSkipFirstLineAndTrimAll(string value)
        {
            if (value == null)
                throw new ArgumentNullException("value");

            return value.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Skip(1).Select(v => v.Trim());
        }
    }
}
