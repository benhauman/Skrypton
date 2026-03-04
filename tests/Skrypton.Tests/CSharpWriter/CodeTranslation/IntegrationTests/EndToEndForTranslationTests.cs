using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//#using Xunit#;

namespace Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests
{
    [TestClass]
    public class EndToEndForTranslationTests : TestBase
    {
        [TestMethod, MyFact]
        public void AscendingLoopWithImplicitStep()
        {
            var source = @"
				Dim i: For i = 1 To 5
				Next
			";
            var expected = new[]
            {
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A loop that exceeds the range of the VBScript "Integer" will result in the loop variable being set to a larger type so that it can describe all of the
        /// values within the loop. Note that there is special handling to identify the case when all loop constraints are constants within VBScript's "Integer"
        /// range, which is why the test AscendingLoopWithImplicitStep does not require an addition "loopStart" variable. That shortcut is not in play here
        /// and so a "loopStart" variable IS required (to determine what type to use to cover the range from (Int16)1 to 32768 - which is implicitly an
        /// Int32 (aka "int") when compiled as C#.
        /// </summary>
        [TestMethod, MyFact]
        public void AscendingLoopThatRollsOverLoopVariableIntoLongType()
        {
            var source = @"
				Dim i: For i = 1 To 32768
				Next
			";
            var expected = new[]
            {
                "var loopStart = _.NUM((Int16)1, 32768);",
                "for (_outer.i = loopStart; _.StrictLTE(_outer.i, 32768); _outer.i = _.ADD(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop range is in the opposite direction to step then it will never be entered in VBScript and so there's no pointing emitting any C# code (this
        /// can only be done if the loop start, end and step are known at compile time - here the start and end are numeric and the loop is implicitly one)
        /// </summary>
        [TestMethod, MyFact]
        public void DescendingLoopWithoutExplicitStepIsOptimisedOut()
        {
            var source = @"
				Dim i: For i = 5 To 1
				Next
			";
            myAssert.AreEqual(
                [],
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void DescendingLoopWithExplicitNegativeStep()
        {
            var source = @"
				Dim i: For i = 5 To 1 Step -1
				Next
			";
            var expected = new[]
            {
                "for (_outer.i = (Int16)5; _.StrictGTE(_outer.i, 1); _outer.i = _.SUBT(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A fractional step on an otherwise small integer range changes the loop variable from being a VBScript "Integer" to a "Double"
        /// </summary>
        [TestMethod, MyFact]
        public void DescendingLoopWithExplicitFractionalStep()
        {
            var source = @"
				Dim i: For i = 1 To 5 Step 0.1
				Next
			";
            var expected = new[]
            {
                "var loopStart = _.NUM((Int16)1, (Int16)5, 0.1);",
                "for (_outer.i = loopStart; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, 0.1))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected,
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void ZeroStepResultsInInfiniteLoopWhenAscending()
        {
            var source = @"
				Dim i: For i = 1 To 5 Step 0
				Next
			";
            var expected = new[]
            {
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5);)",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop has fixed contraints that indicate a negative direction and a zero step, the loop will not be entered and can be optimised out
        /// </summary>
        [TestMethod, MyFact]
        public void ZeroStepIsOptimisedOutForDescendingLoop()
        {
            var source = @"
				Dim i: For i = 5 To 1 Step 0
				Next
			";
            myAssert.AreEqual(
                [],
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop has fixed contraints that indicate a negative direction and a zero step, the loop will not be entered and can be optimised out
        /// </summary>
        [TestMethod, MyFact]
        public void FixedNegativeStepResultsInLoopBeingOptimisedOutIfItIsFixedAndPositive()
        {
            var source = @"
				Dim i: For i = 1 To 5 Step - 1
				Next
			";
            myAssert.AreEqual(
                [],
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a loop is known at compile time to run in a negative direction and no step is specified then the loop is never entered and can be optimised out
        /// </summary>
        [TestMethod, MyFact]
        public void FixedNegativeLoopWithoutExplicitStepIsOptimisedOut()
        {
            var source = @"
				Dim i: For i = 5 To 1
				Next
			";
            myAssert.AreEqual(
                [],
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void FixedAscendingLoopWithExplicitPositiveStep()
        {
            var source = @"
				Dim i: For i = 1 To 5 Step 2
				Next
			";
            var expected = new[]
            {
                "for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 5); _outer.i = _.ADD(_outer.i, (Int16)2))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        [TestMethod, MyFact]
        public void FixedDescendingLoopWithExplicitNegativeStep()
        {
            var source = @"
				Dim i: For i = 5 To 1 Step -1
				Next
			";
            var expected = new[]
            {
                "for (_outer.i = (Int16)5; _.StrictGTE(_outer.i, 1); _outer.i = _.SUBT(_outer.i, (Int16)1))",
                "{",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop start, end and step values are not known until runtime then their values must be determined once and then applied to a loop (in
        /// VBScript, the constraints are not re-evaluated each loop iteration). The loop may only be entered if there is a zero or positive step and a
        /// non-descending loop or if there is a negative step and a descending loop. Similarly, the termination condition operator may be a less-than-
        /// or-equal-to comparison or a greater-than-or-equal-to, depending upon loop direction.
        /// </summary>
        [TestMethod, MyFact]
        public void RuntimeVariableLoopBoundariesAndStep()
        {
            var source = @"
				For i = a To b Step c
				Next
			";
            var expected = new[]
            {
                "var loopEnd = _.NUM(_env.b);",
                "var loopStep = _.NUM(_env.c);",
                "var loopStart = _.NUM(_env.a, loopEnd, loopStep);",
                "if ((_.StrictLTE(loopStart, loopEnd) && _.StrictGTE(loopStep, 0))",
                "|| (_.StrictGT(loopStart, loopEnd) && _.StrictLT(loopStep, 0)))",
                "{",
                "for (_env.i = loopStart;",
                "    (_.StrictGTE(loopStep, 0) && _.StrictLTE(_env.i, loopEnd)) || (_.StrictLT(loopStep, 0) && _.StrictGTE(_env.i, loopEnd));",
                "     _env.i = _.ADD(_env.i, loopStep))",
                "{",
                "}",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If there are non-compile-time-known-numeric-constant constraints and error-trapping may be enabled, then the constraints must be evaluated
        /// first. If these constraints are successfully evaluated, then the loop proceeds as would be expected, but there must be error-trapping around
        /// the loop itself, so that if the termination condition or loop-variable-addition/subtraction fails then the loop will terminate. There must
        /// also be error-trapping around each statement within the loop, so that if any one of them fails then the others may still be processed (if
        /// the error-trapping token is enabled at that point during the runtime execution). HOWEVER, the craziest bit is that if evaluation of any
        /// of the loop constraints fails then no further constraint evaluation will be attempted (they are processed in the order of From, To and
        /// Step) but the loop WILL be executed once. For this iteration, the loop variable will not be altered (so it will be left as null in
        /// this example, but if it had been set to "a" before the loop then it would remain set to "a") and neither the loop termination
        /// condition nor the increment work will be attempted.
        /// </summary>
        [TestMethod, MyFact]
        public void RuntimeVariableLoopBoundariesWithErrorTrapping()
        {
            var source = @"
				On Error Resume Next
				For i = a To b
					WScript.Echo i
				Next
			";
            var expected = new[]
            {
                "var errOn = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);",
                "object loopEnd = 0, loopStart = 0;",
                "var loopConstraintsInitialized = false;",
                "_.HANDLEERROR(errOn, () => {",
                "   loopEnd = _.NUM(_env.b);",
                "   loopStart = _.NUM(_env.a);",
                "   if ((loopStart is DateTime) || (loopStart is Decimal))",
                "       _env.i = loopStart;",
                "   loopStart = _.NUM(_env.a, loopEnd, (Int16)1);",
                "   loopConstraintsInitialized = true;",
                "});",
                "if (_.StrictLTE(loopStart, loopEnd))",
                "{",
                "   if (loopConstraintsInitialized)",
                "       _env.i = loopStart;",
                "   while (true)",
                "   {",
                "       _.HANDLEERROR(errOn, () => {",
                "           _.CALL(this, _env.WScript, \"Echo\", _.ARGS.Ref(_env.i, v => { _env.i = v; }));",
                "       });",
                "       if (!loopConstraintsInitialized)",
                "           break;",
                "       var continueLoop = false;",
                "       _.HANDLEERROR(errOn, () => {",
                "           _env.i = _.ADD(_env.i, (Int16)1);",
                "           continueLoop = _.StrictLTE(_env.i, loopEnd);",
                "       });",
                "       if (!continueLoop)",
                "           break;",
                "   }",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn);"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If the loop constraints are known numeric values at translation time then enabling error-handling is relatively easy. The loop needs to be
        /// wrapped in error-trapping in case the termination condition or loop variable addition/subtraction fail, then the individual statements
        /// within the loop need wrapping as well. But without any dynamic loop constraints to be evaluated, it's a lot simpler - no evaluation
        /// of contraints to trap or guard clause around the loop to worry about.
        /// </summary>
        [TestMethod, MyFact]
        public void AscendingLoopWithImplicitStepAndErrorTrappingEnabled()
        {
            var source = @"
				On Error Resume Next
				For i = 1 To 10
					WScript.Echo i
				Next
			";
            var expected = new[]
            {
                "var errOn = _.GETERRORTRAPPINGTOKEN();",
                "_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);",
                "_env.i = (Int16)1;",
                "while (true)",
                "{",
                "   _.HANDLEERROR(errOn, () => {",
                "       _.CALL(this, _env.WScript, \"Echo\", _.ARGS.Ref(_env.i, v => { _env.i = v; }));",
                "   });",
                "   var continueLoop = false;",
                "   _.HANDLEERROR(errOn, () => {",
                "       _env.i = _.ADD(_env.i, (Int16)1);",
                "       continueLoop = _.StrictLTE(_env.i, 10);",
                "   });",
                "   if (!continueLoop)",
                "       break;",
                "}",
                "_.RELEASEERRORTRAPPINGTOKEN(errOn);"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// A loop variable may be of type "Byte" but only if the start, end and step values are all of type "Byte" - if there is no step explicitly
        /// specified then the default "Integer" 1 will be used and so the loop variable will become type "Integer" (this test doesn't really show
        /// this completely since the translated code is not executed and it would depend upon the support class implementation but it seemed like
        /// it was worth recording here to make the point, also see the NUM test "BytesWithAnInteger")
        /// </summary>
        [TestMethod, MyFact]
        public void ByteLoopStartAndEndValuesWithImplicitStepWillGetAnIntegerStep()
        {
            var source = @"
				Dim i: For i = CByte(1) To CByte(5)
				Next
			";
            var expected = new[]
            {
                "var loopEnd = _.CBYTE(5);",
                "var loopStart = _.NUM(_.CBYTE(1), loopEnd, (Int16)1);",
                "if (_.StrictLTE(loopStart, loopEnd))",
                "{",
                "    for (_outer.i = loopStart; _.StrictLTE(_outer.i, loopEnd); _outer.i = _.ADD(_outer.i, (Int16)1))",
                "    {",
                "    }",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is the complement to ByteLoopStartAndEndValuesWithImplicitStepWillGetAnIntegerStep, it illustrates how a loop would be constructed
        /// in order to have the loop variable be of type "Byte".
        /// </summary>
        [TestMethod, MyFact]
        public void ByteLoopStartAndEndAndStepValuesWillGetByteLoopVariable()
        {
            var source = @"
				Dim i: For i = CByte(1) To CByte(5) Step CByte(1)
				Next
			";
            var expected = new[]
            {
                "var loopEnd = _.CBYTE(5);",
                "var loopStep = _.CBYTE(1);",
                "var loopStart = _.NUM(_.CBYTE(1), loopEnd, loopStep);",
                "if ((_.StrictLTE(loopStart, loopEnd) && _.StrictGTE(loopStep, 0))",
                "|| (_.StrictGT(loopStart, loopEnd) && _.StrictLT(loopStep, 0)))",
                "{",
                "    for (_outer.i = loopStart;",
                "        (_.StrictGTE(loopStep, 0) && _.StrictLTE(_outer.i, loopEnd)) || (_.StrictLT(loopStep, 0) && _.StrictGTE(_outer.i, loopEnd));",
                "         _outer.i = _.ADD(_outer.i, loopStep))",
                "    {",
                "    }",
                "}"
            };
            //base.TestCSharpCodeTranslationWithoutScaffoldingTranslator(source, expected);
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        // TODO: Various variable-ascending/descending/step combinations

        /// <summary>
        /// When the translation of a for loop is completed, any undeclared variables should not be flushed (declared) within the scope of the loop, it
        /// must be within the scope-defining parent. If within a function then these must be local variables. This test also covers a fix where the loop
        /// variable was not getting identified as an undeclared variable when it should have been.
        /// </summary>
        [TestMethod, MyFact]
        public void UndeclaredVariablesShouldNotBeFlushedAtForBlockEnd()
        {
            var source = @"
				Function F1
					For i = 1 To 5
						WScript.Echo j
					Next
				End Function
			";
            var expected = new[]
            {
                "public object F1()",
                "{",
                "    object F1_retVal = null;",
                "    object j = null; /* Undeclared in source */",
                "    object i = null; /* Undeclared in source */",
                "    for (i = (Int16)1; _.StrictLTE(i, 5); i = _.ADD(i, (Int16)1))",
                "    {",
                "        _.CALL(this, _env.WScript, \"Echo\", _.ARGS.Ref(j, v => { j = v; }));",
                "    }",
                "    return F1_retVal;",
                "}"
            };
            myAssert.AreEqual(
                expected.Select(s => s.Trim()).ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a FOR loop exists within a function F1 and one of F1's arguments is used to determine a loop constraint and that argument is passed to F1 ByRef and that argument is passed
        /// to another function while determining the loop constraints then a ByRef mapping will be required for the F1 argument. This is because the argument will be referenced inside a
        /// lambda when passed as Ref argument and it is not legal C# to reference a ref argument within a lambda.
        /// </summary>
        [TestMethod, MyFact]
        public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired()
        {
            var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
				End Function";

            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					object i = null;
					object loopEnd = 0, loopStart = 0;
					var loopConstraintsInitialized = false;
					object byrefalias = x;
					try
					{
						loopEnd = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias, v => { byrefalias = v; })));
						loopStart = _.NUM((Int16)1);
						if ((loopStart is DateTime) || (loopStart is Decimal))
							i = loopStart;
						loopStart = _.NUM((Int16)1, loopEnd);
						loopConstraintsInitialized = true;
					}
					finally { x = byrefalias; }
					if (_.StrictLTE(loopStart, loopEnd))
					{
						for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return F1_retVal;
				}

				public object F2(ref object value)
				{
					return _.VAL(value);
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This test is a companion of IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired and illustrates that a ByRef mapping is not required
        /// for the F1 argument if the argument is passed in ByVal (while the argument still needs to be referenced in a lambda when passed to F2 as a ByRef argument, it's not a ref argument in
        /// F1 and so we don't need to jump through any hoops to avoid illegal C#)
        /// </summary>
        [TestMethod, MyFact]
        public void IfByValArgumentIsRequiredForLoopConstraintAndIsPassedToAnotherFunctionByRefThenNoByRefMappingIsRequiredAsTheFirstArgumentWasByVal()
        {
            var source = @"
				Function F1(ByVal x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
				End Function";

            var expected = @"
				public object F1(object x)
				{
					object F1_retVal = null;
					object i = null;
					var loopEnd = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(x, v => { x = v; })));
					var loopStart = _.NUM((Int16)1, loopEnd);
					if (_.StrictLTE(loopStart, loopEnd))
					{
						for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return F1_retVal;
				}

				public object F2(ref object value)
				{
					return _.VAL(value);
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a companion to IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionByRefThenByRefMappingRequired and illustrates a limitation of the translation process.
        /// The function F1 takes a ByRef argument which is then passed to F2 as the loop constraints are initialised. Although F2 accepts the argument ByVal, the translation analysis does not
        /// go deeply enough to realise this and presumes that F2 may take the argument ByRef - as such, it tries to pass it ByRef (just in case) and so needs to reference the F1 argument within
        /// a lambda, which would not be legal C# and so a ByRef mapping is unfortunately required.
        /// </summary>
        [TestMethod, MyFact]
        public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionThenByRefMappingRequired()
        {
            var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByVal value)
					F2 = value
				End Function";

            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					object i = null;
					object loopEnd = 0, loopStart = 0;
					var loopConstraintsInitialized = false;
					object byrefalias = x;
					try
					{
						loopEnd = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias, v => { byrefalias = v; })));
						loopStart = _.NUM((Int16)1);
						if ((loopStart is DateTime) || (loopStart is Decimal))
							i = loopStart;
						loopStart = _.NUM((Int16)1, loopEnd);
						loopConstraintsInitialized = true;
					}
					finally { x = byrefalias; }
					if (_.StrictLTE(loopStart, loopEnd))
					{
						for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return F1_retVal;
				}

				public object F2(object value)
				{
					return _.VAL(value);
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// This is a companion to IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToAnotherFunctionThenByRefMappingRequired and shows that we can make things a little better by
        /// presuming that all built-in functions take arguments ByVal (which I'm fairly confident is always the case), which means that ByRef mappings may be avoided for some cases.
        /// </summary>
        [TestMethod, MyFact]
        public void IfByRefArgumentIsRequiredForLoopConstraintsAndIsPassedToBuiltInFunctionByRefThenNoByRefMappingRequired()
        {
            var source = @"
				Function F1(ByRef x)
					Dim i: For i = 1 To UBOUND(x)
					Next
				End Function";

            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					object i = null;
					var loopEnd = _.UBOUND(x);
					var loopStart = _.NUM((Int16)1, loopEnd);
					if (_.StrictLTE(loopStart, loopEnd))
					{
						for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
						{
						}
					}
					return F1_retVal;
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a ByRef argument is passed to a method F1 that uses the argument when evaluating loop constraints within an error-trapping block then a ByRef mapping will be required because the
        /// ByRef argument will need to be accessed within a lambda (inside the HANDLEERROR block), which is not legal in C#.
        /// </summary>
        [TestMethod, MyFact]
        public void IfByRefArgumentIsRequiredForKnownLoopConstraintsAndLoopWrappedInErrorTrappingThenByRefMappingRequired()
        {
            var source = @"
				Function F1(ByRef x)
					On Error Resume Next
					Dim i: For i = 1 To F2(x)
					Next
				End Function

				Function F2(ByRef value)
					F2 = value
					value = 123
				End Function";

            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					object i = null;
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					object loopEnd = 0, loopStart = 0;
					var loopConstraintsInitialized = false;
					object byrefalias = x;
					try
					{
						_.HANDLEERROR(errOn, () => {
							loopEnd = _.NUM(_.CALL(this, _outer, ""F2"", _.ARGS.Ref(byrefalias, v => { byrefalias = v; })));
							loopStart = _.NUM((Int16)1);
							if ((loopStart is DateTime) || (loopStart is Decimal))
								i = loopStart;
							loopStart = _.NUM((Int16)1, loopEnd);
							loopConstraintsInitialized = true;
						});
					}
					finally { x = byrefalias; }
					if (_.StrictLTE(loopStart, loopEnd))
					{
						if (loopConstraintsInitialized)
							i = loopStart;
						while (true)
						{
							if (!loopConstraintsInitialized)
								break;
							var continueLoop = false;
							_.HANDLEERROR(errOn, () => {
								i = _.ADD(i, (Int16)1);
								continueLoop = _.StrictLTE(i, loopEnd);
							});
							if (!continueLoop)
								break;
						}
					}
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}

				public object F2(ref object value)
				{
					object F2_retVal = null;
					F2_retVal = _.VAL(value);
					value = (Int16)123;
					return F2_retVal;
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }

        /// <summary>
        /// If a ByRef argument is passed to a method F1 that uses the argument when evaluating loop constraints within an error-trapping block then a ByRef mapping will be required because the
        /// ByRef argument will need to be accessed within a lambda (inside the HANDLEERROR block), which is not legal in C#. When it is known that the ByRef argument will not be changed by the
        /// loop constraint evaluation, the ByRef mapping is readonly; meaning that no try..finally wrapping is required to write the byref-temp-value back over the method argument (because it
        /// is known that the temporary value will not have been manipulated).
        /// </summary>
        [TestMethod, MyFact]
        public void IfByRefArgumentIsRequiredForKnownReadOnlyLoopConstraintsAndLoopWrappedInErrorTrappingThenReadOnlyByRefMappingRequired()
        {
            var source = @"
				Function F1(ByRef x)
					On Error Resume Next
					Dim i: For i = 1 To x + 1
					Next
				End Function";

            var expected = @"
				public object F1(ref object x)
				{
					object F1_retVal = null;
					var errOn = _.GETERRORTRAPPINGTOKEN();
					object i = null;
					_.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
					object loopEnd = 0, loopStart = 0;
					var loopConstraintsInitialized = false;
					object byrefalias = x;
					_.HANDLEERROR(errOn, () => {
						loopEnd = _.NUM(_.ADD(byrefalias, (Int16)1));
						loopStart = _.NUM((Int16)1);
						if ((loopStart is DateTime) || (loopStart is Decimal))
							i = loopStart;
						loopStart = _.NUM((Int16)1, loopEnd);
						loopConstraintsInitialized = true;
					});
					if (_.StrictLTE(loopStart, loopEnd))
					{
						if (loopConstraintsInitialized)
							i = loopStart;
						while (true)
						{
							if (!loopConstraintsInitialized)
								break;
							var continueLoop = false;
							_.HANDLEERROR(errOn, () => {
								i = _.ADD(i, (Int16)1);
								continueLoop = _.StrictLTE(i, loopEnd);
							});
							if (!continueLoop)
								break;
						}
					}
					_.RELEASEERRORTRAPPINGTOKEN(errOn);
					return F1_retVal;
				}";

            myAssert.AreEqual(
                expected.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray(),
                WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, source, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies)
            );
        }
    }
}
