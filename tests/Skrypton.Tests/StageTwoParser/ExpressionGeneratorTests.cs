
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.StageTwoParser.ExpressionParsing;
using Skrypton.StageTwoParser.Tokens;
using Skrypton.Tests.Shared.Comparers;
//#using Xunit#;

namespace Skrypton.Tests.StageTwoParser
{
    [TestClass]
	public class ExpressionGeneratorTests
	{
		[TestMethod, MyFact]
		public void DirectFunctionCallWithNoArgumentsAndNoBrackets()
		{
            // Test
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("Test", 0))
					)
				},
				ExpressionGenerator.Generate(
					new[] {
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void DirectFunctionCallWithNoArgumentsWithBrackets()
		{
            // Test()
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("Test", 0), CallExpressionSegment.ArgumentBracketPresenceOptions.Present)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("Test", 0),
						new OpenBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ObjectFunctionCallWithNoArgumentsAndNoBrackets()
		{
            // a.Test
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("a", 0), new NameToken("Test", 0) }
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void NestedObjectFunctionCallWithNoArgumentsAndNoBrackets()
		{
			// a.b.Test
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("a", 0), new NameToken("b", 0), new NameToken("Test", 0) }
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new MemberAccessorToken(0),
						new NameToken("b", 0),
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void DirectFunctionCallWithOneArgument()
		{
            // Test(1)
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("Test", 0) },
							new[] { new NumericValueToken("1", 0) }
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("Test", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void DirectFunctionCallWithTwoArguments()
		{
            // Test(1, 2)
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("Test", 0) },
							new[] { new NumericValueToken("1", 0) },
							new[] { new NumericValueToken("2",0) }
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("Test", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new ArgumentSeparatorToken(0),
						new NumericValueToken("2",0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void DirectFunctionCallWithTwoArgumentsOneIsNestedDirectionFunctionCallWithOneArgument()
		{
            // Test(Test2(1), 2)
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("Test", 0) },
							EXP(
								CALL(
									new[] { new NameToken("Test2", 0) },
									new[] { new NumericValueToken("1", 0) }
								)
							),
							EXP(CALL(new NumericValueToken("2",0)))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("Test", 0),
						new OpenBrace(0),
						new NameToken("Test2", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new CloseBrace(0),
						new ArgumentSeparatorToken(0),
						new NumericValueToken("2",0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ArrayElementFunctionCallWithNoArguments()
		{
			// a(0).Test
			myAssert.AreEqual(new[]
				{
					EXP(
						CALLSET(
							CALL(
								new[] { new NameToken("a", 0) },
								new[] { new NumericValueToken("0", 0) }
							),
							CALL(
								new[] { new NameToken("Test", 0) }
							)
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ObjectPropertyArrayElementFunctionCallWithNoArguments()
		{
			// a.b(0).Test
			myAssert.AreEqual(new[]
				{
					EXP(
						CALLSET(
							CALL(
								new[] { new NameToken("a", 0), new NameToken("b", 0) },
								new[] { new NumericValueToken("0", 0) }
							),
							CALL(
								new[] { new NameToken("Test", 0) }
							)
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new MemberAccessorToken(0),
						new NameToken("b", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ArrayElementNestedFunctionCallWithNoArguments()
		{
			// a(0).b.Test
			myAssert.AreEqual(new[]
				{
					EXP(
						CALLSET(
							CALL(
								new[] { new NameToken("a", 0) },
								new[] { new NumericValueToken("0", 0) }
							),
							CALL(
								new[] { new NameToken("b", 0), new NameToken("Test", 0) }
							)
						)
					)
				},
				ExpressionGenerator.Generate(new IToken[] {
						new NameToken("a", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new MemberAccessorToken(0),
						new NameToken("b", 0),
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void JaggedArrayAccess()
		{
			// a(0)(1)
			myAssert.AreEqual(new[]
				{
					EXP(
						CALLSET(
							CALL(
								new[] { new NameToken("a", 0) },
								new[] { new NumericValueToken("0", 0) }
							),
							CALLARGSONLY(
								new[] { new NumericValueToken("1", 0) }
							)
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new CloseBrace(0),
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// Additional brackets will be applied around all operations to ensure that VBScript operator rules are always maintained (if the operators
		/// are all equivalent in terms of priority, terms will be bracketed from left-to-right, so a and b should be bracketed together)
		/// </summary>
		[TestMethod, MyFact]
		public void AdditionWithThreeTerms()
		{
			// a + b + c
			myAssert.AreEqual(new[]
				{
					EXP(
						BR(
							CALL(new NameToken("a", 0)),
							OP(new OperatorToken("+", 0)),
							CALL(new NameToken("b", 0))
						),
						OP(new OperatorToken("+", 0)),
						CALL(new NameToken("c", 0))
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new NameToken("b", 0),
						new OperatorToken("+", 0),
						new NameToken("c", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// Multiplication should take precedence over addition so b and c should be bracketed together
		/// </summary>
		[TestMethod, MyFact]
		public void AdditionAndMultiplicationWithThreeTerms()
		{
			// a + b * c
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new OperatorToken("+", 0)),
						BR(
							CALL(new NameToken("b", 0)),
							OP(new OperatorToken("*", 0)),
							CALL(new NameToken("c", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void AdditionAndMultiplicationWithThreeTermsWhereTheThirdTermIsAnArrayElement()
		{
			// a + b * c(0)
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new OperatorToken("+", 0)),
						BR(
							CALL(new NameToken("b", 0)),
							OP(new OperatorToken("*", 0)),
							CALL(
								new[] { new NameToken("c", 0) },
								new[] { new NumericValueToken("0", 0) }
							)
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// This will try to ensure that the bracket around the array access doesn't interfere with the formatting of the fourth term
		/// </summary>
		[TestMethod, MyFact]
		public void AdditionAndMultiplicationAndAdditionWithFourTermsWhereTheThirdTermIsAnArrayElement()
		{
			// a + b * c(0) + d
			myAssert.AreEqual(new[]
				{
					EXP(
						BR(
							CALL(new NameToken("a", 0)),
							OP(new OperatorToken("+", 0)),
							BR(
								CALL(new NameToken("b", 0)),
								OP(new OperatorToken("*", 0)),
								CALL(
									new[] { new NameToken("c", 0) },
									new[] { new NumericValueToken("0", 0) }
								)
							)
						),
						OP(new OperatorToken("+", 0)),
						CALL(new NameToken("d", 0))
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new OperatorToken("+", 0),
						new NameToken("d", 0),
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// If an operation is already bracketed then additional brackets should not be added around the operation, they would be unnecessary
		/// </summary>
		[TestMethod, MyFact]
		public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketing()
		{
			// a + (b * c)
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new OperatorToken("+", 0)),
						BR(
							CALL(new NameToken("b", 0)),
							OP(new OperatorToken("*", 0)),
							CALL(new NameToken("c", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new OpenBrace(0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void AlreadyBracketedOperationsShouldNotGetUnnecessaryBracketingIfTheyAppearInTheMiddleOfTheExpression()
		{
			// a + (b * c) + d
			myAssert.AreEqual(new[]
				{
					EXP(
						BR(
							CALL(new NameToken("a", 0)),
							OP(new OperatorToken("+", 0)),
							BR(
								CALL(new NameToken("b", 0)),
								OP(new OperatorToken("*", 0)),
								CALL(new NameToken("c", 0))
							)
						),
						OP(new OperatorToken("+", 0)),
						CALL(new NameToken("d", 0))
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new OpenBrace(0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0),
						new CloseBrace(0),
						new OperatorToken("+", 0),
						new NameToken("d", 0),
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// Arithmetic operations should take precedence over comparisons so b and c should be bracketed together
		/// </summary>
		[TestMethod, MyFact]
		public void AdditionAndEqualityComparisonWithThreeTerms()
		{
			// a = b + c
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new ComparisonOperatorToken("=", 0)),
						BR(
							CALL(new NameToken("b", 0)),
							OP(new OperatorToken("+", 0)),
							CALL(new NameToken("c", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new ComparisonOperatorToken("=", 0),
						new NameToken("b", 0),
						new OperatorToken("+", 0),
						new NameToken("c", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// This covers an array of different types of expression
		/// </summary>
		[TestMethod, MyFact]
		public void TestArrayAccessObjectAccessMethodArgumentsMixedArithmeticAndComparisonOperations()
		{
			// a + b * c.d(Test(0), 1) + e = f
			myAssert.AreEqual(new[]
				{
					EXP(
						BR(
							BR(
								CALL(new NameToken("a", 0)),
								OP(new OperatorToken("+", 0)),
								BR(
									CALL(new NameToken("b", 0)),
									OP(new OperatorToken("*", 0)),
									CALL(
										new[] { new NameToken("c", 0), new NameToken("d", 0) },
										EXP(
											CALL(
												new[] { new NameToken("Test", 0) },
												new[] { new NumericValueToken("0", 0) }
											)
										),
										EXP(CALL(new NumericValueToken("1", 0)))
									)
								)
							),
							OP(new OperatorToken("+", 0)),
							CALL(new NameToken("e", 0))
						),
						OP(new ComparisonOperatorToken("=", 0)),
						CALL(new NameToken("f", 0))
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("+", 0),
						new NameToken("b", 0),
						new OperatorToken("*", 0),
						new NameToken("c", 0),
						new MemberAccessorToken(0),
						new NameToken("d", 0),
						new OpenBrace(0),
						new NameToken("Test", 0),
						new OpenBrace(0),
						new NumericValueToken("0", 0),
						new CloseBrace(0),
						new ArgumentSeparatorToken(0),
						new NumericValueToken("1", 0),
						new CloseBrace(0),
						new OperatorToken("+", 0),
						new NameToken("e", 0),
						new ComparisonOperatorToken("=", 0),
						new NameToken("f", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// To make it clear that the "-" is a one-sided operation (a negation, not a subtraction), it should be bracketed
		/// </summary>
		[TestMethod, MyFact]
		public void NegatedTermsShouldBeBracketed()
		{
			// a * -b
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new OperatorToken("*", 0)),
						BR(
							OP(new OperatorToken("-", 0)),
							CALL(new NameToken("b", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new OperatorToken("*", 0),
						new OperatorToken("-", 0),
						new NameToken("b", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// This is the boolean equivalent of NegatedTermsShouldBeBracketed
		/// </summary>
		[TestMethod, MyFact]
		public void LogicalInversionsTermsShouldBeBracketed()
		{
			// a AND NOT b
			myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new LogicalOperatorToken("AND", 0)),
						BR(
							OP(new LogicalOperatorToken("NOT", 0)),
							CALL(new NameToken("b", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new LogicalOperatorToken("AND", 0),
						new LogicalOperatorToken("NOT", 0),
						new NameToken("b", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// This exercises a fix for the translation of "NOT NOT a", which was bracketing the two NOTs together instead of (NOT(NOT(a))
		/// </summary>
		[TestMethod, MyFact]
		public void AdjacentLogicalInversionsShouldBracketWithOtherTermsAndNotEachOther()
		{
            // a AND NOT NOT b
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(new NameToken("a", 0)),
						OP(new LogicalOperatorToken("AND", 0)),
						BR(
							OP(new LogicalOperatorToken("NOT", 0)),
							BR(
								OP(new LogicalOperatorToken("NOT", 0)),
								CALL(new NameToken("b", 0))
							)
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("a", 0),
						new LogicalOperatorToken("AND", 0),
						new LogicalOperatorToken("NOT", 0),
						new LogicalOperatorToken("NOT", 0),
						new NameToken("b", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// This indicates different precedence that is applied to a NOT operation depending upon content, as compared to the test
		/// LogicalInversionsTermsShouldBeBracketed
		/// </summary>
		[TestMethod, MyFact]
		public void NegationOperationHasLessPrecendenceThanComparsionOperations()
		{
            // NOT a IS Nothing
            myAssert.AreEqual(new[]
				{
					EXP(
						OP(new LogicalOperatorToken("NOT", 0)),
						BR(
							CALL(new NameToken("a", 0)),
							OP(new ComparisonOperatorToken("IS", 0)),
							new BuiltInValueExpressionSegment(new BuiltInValueToken("Nothing", 0))
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new LogicalOperatorToken("NOT", 0),
						new NameToken("a", 0),
						new ComparisonOperatorToken("IS", 0),
						new BuiltInValueToken("Nothing", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void NewInstanceRequestsShouldNotBeConfusedWithCallExpressions()
		{
            // new Test
            myAssert.AreEqual(new[]
				{
					EXP(
						NEW("Test", 0)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new KeyWordToken("new", 0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		/// <summary>
		/// If a function (or property) argument is wrapped in brackets then it should be passed ByVal even when otherwise it would be passed ByRef.
		/// This means that brackets can have special significance and should not be removed, even from places where they would have significance or
		/// meaning in C#.
		/// </summary>
		[TestMethod, MyFact]
		public void BracketsShouldNotBeRemovedFromSingleArgumentCallStatements()
		{
            // CALL Test((a))
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new NameToken("Test", 0),
							EXP(
								BR(CALL(new NameToken("a", 0)))
							)
						)
					)
				},
				ExpressionGenerator.Generate(
						new IToken[] {
						new NameToken("Test", 0),
						new OpenBrace(0),
						new OpenBrace(0),
						new NameToken("a", 0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ObjectFunctionCallWithNoArgumentsAndNoBracketsThatReliesUponDirectedWithReference()
		{
            // ".Test" within "WITH a"
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new DoNotRenameNameToken("a", 0), new NameToken("Test", 0) }
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new MemberAccessorToken(0),
						new NameToken("Test", 0)
					},
					directedWithReferenceIfAny: new DoNotRenameNameToken("a", 0),
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void PropertyAccessOnNumberLiteralResultsInException()
		{
            // "WScript.Echo 1.a" results in a compile time error from the VBScript parser
            // Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
            // and so we need to insert brackets around the "1.a" argument even though they would not necessarily be present in the source code
            myAssert.Throws<ArgumentException>(() =>
			{
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new MemberAccessorToken(0),
						new NameToken("a", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				);
			});
		}

		[TestMethod, MyFact]
		public void NumericLiteralPropertyAccessResultsInException()
		{
            // "WScript.Echo a.1" results in a compile time error from the VBScript parser
            // Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
            // and so we need to insert brackets around the "a.1" argument even though they would not necessarily be present in the source code
            myAssert.Throws<ArgumentException>(() =>
			{
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NameToken("a", 0),
						new MemberAccessorToken(0),
						new NumericValueToken("1", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				);
			});
		}

		[TestMethod, MyFact]
		public void ZeroArgumentMethodAccessOnNumberLiteralResultsInException()
		{
            // "WScript.Echo 1.a()" results in a compile time error from the VBScript parser
            // Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
            // and so we need to insert brackets around the "1.a()" argument even though they would not necessarily be present in the source code
            myAssert.Throws<ArgumentException>(() =>
			{
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new MemberAccessorToken(0),
						new NameToken("a", 0),
						new OpenBrace(0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				);
			});
		}

		[TestMethod, MyFact]
		public void ZeroArgumentDefaultMethodAccessOnNumberLiteralResultsInRuntimeError()
		{
			// "WScript.Echo 1()" results in a runtime error ("Type mismatch")
			// Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
			// and so we need to insert brackets around the "1()" argument even though they would not necessarily be present in the source code
			var runtimeErrorExpressionSegment = new RuntimeErrorExpressionSegment(
				"1()",
				new IToken[] { new NumericValueToken("1", 0), new OpenBrace(0), new CloseBrace(0) },
				typeof(TypeMismatchException),
				"'[number: 1]' is called like a function"
			);
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("wscript", 0), new NameToken("echo", 0) },
							new Expression(new[] { runtimeErrorExpressionSegment })
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new OpenBrace(0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: new DoNotRenameNameToken("a", 0),
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void SingleArgumentMethodAccessOnNumberLiteralResultsInException()
		{
            // "WScript.Echo 1.a(b)" results in a compile time error from the VBScript parser
            // Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
            // and so we need to insert brackets around the "1.a(b)" argument even though they would not necessarily be present in the source code
           myAssert.Throws<ArgumentException>(() =>
			{
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new MemberAccessorToken(0),
						new NameToken("a", 0),
						new OpenBrace(0),
						new NameToken("b", 0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				);
			});
		}

		[TestMethod, MyFact]
		public void PropertyAccessOnStringLiteralResultsInRuntimeError()
		{
            // "WScript.Echo \"1\".a" results in a runtime "Object required" runtime error. HOWEVER, this is handled at runtime by the CALL implementation,
            // the "\"1\".a" attempt should be translated into _.CALL("1", "a"), which should fail at evaluation
            // Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
            // and so we need to insert brackets around the "\"1\".a" argument even though they would not necessarily be present in the source code
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("wscript", 0), new NameToken("echo", 0) },
							new Expression(new[] {
								CALL(new IToken[] { new StringToken("1", 0), new NameToken("a", 0) })
							})
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new StringToken("1", 0),
						new MemberAccessorToken(0),
						new NameToken("a", 0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: new DoNotRenameNameToken("a", 0),
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void ZeroArgumentMethodAccessOnStringLiteralResultsInRuntimeError()
		{
			// "WScript.Echo \"1\".a()" results in a runtime "Object required" runtime error
			// Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
			// and so we need to insert brackets around the "\"1\".a()" argument even though they would not necessarily be present in the source code
			var runtimeErrorExpressionSegment = new RuntimeErrorExpressionSegment(
				"1()",
				new IToken[] { new NumericValueToken("1", 0), new OpenBrace(0), new CloseBrace(0) },
				typeof(TypeMismatchException),
				"'[number: 1]' is called like a function"
			);
            myAssert.AreEqual(new[]
				{
					EXP(
						CALL(
							new[] { new NameToken("wscript", 0), new NameToken("echo", 0) },
							new Expression(new[] { runtimeErrorExpressionSegment })
						)
					)
				},
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new OpenBrace(0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: new DoNotRenameNameToken("a", 0),
					warningLogger: warning => { }
				),
				new ExpressionSetComparer()
			);
		}

		[TestMethod, MyFact]
		public void SingleArgumentMethodAccessOnStringLiteralResultsInException()
		{
			// "WScript.Echo \"1\".a(b)" results in a runtime "Object required" runtime error
			// Note: The ExpressionGenerator expects bracketing to be "normalised" on no-value-returning functions (such as the WScript.Echo call)
			// and so we need to insert brackets around the "\1\.a(b)" argument even though they would not necessarily be present in the source code
			myAssert.Throws<ArgumentException>(() =>
			{
				ExpressionGenerator.Generate(
					new IToken[] {
						new NameToken("wscript", 0),
						new MemberAccessorToken(0),
						new NameToken("echo", 0),
						new OpenBrace(0),
						new NumericValueToken("1", 0),
						new MemberAccessorToken(0),
						new NameToken("a", 0),
						new OpenBrace(0),
						new NameToken("b", 0),
						new CloseBrace(0),
						new CloseBrace(0)
					},
					directedWithReferenceIfAny: null,
					warningLogger: warning => { }
				);
			});
		}

		// TODO: Built-in constants and boolean member access attempts (these are consistent with string literals in all cases)

		/// <summary>
		/// Create a BracketedExpressionSegment from a set of expressions
		/// </summary>
		private static BracketedExpressionSegment BR(IEnumerable<IExpressionSegment> segments)
		{
			return new BracketedExpressionSegment(segments);
		}

		/// <summary>
		/// Create a BracketedExpressionSegment from a set of expressions
		/// </summary>
		private static BracketedExpressionSegment BR(params IExpressionSegment[] segments)
		{
			return new BracketedExpressionSegment((IEnumerable<IExpressionSegment>)segments);
		}

		private static CallSetExpressionSegment CALLSET(params IExpressionSegment[] segments)
		{
			return new CallSetExpressionSegment(segments.Cast<CallSetItemExpressionSegment>());
		}

		/// <summary>
		/// This method signature is required by Visual Studio 2015 to remove any ambiguity between calls to CALL which specify an IToken set since it is not clear
		/// whether the signature which takes an IToken set and a params IToken set or the one that takes a params Expression set would be a better match (I'm not
		/// sure why Visual Studio 2013 didn't pick up this ambiguity, but it was new for 2015)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens)
		{
			return CALL(memberAccessTokens, new Expression[0]);
		}

		/// <summary>
		/// Create an CallExpressionSegment from member access tokens and argument expressions (the zeroArgBrackets is only considered if arguments is an empty set,
		/// if arguments is empty and zeroArgBrackets is null then a Absent will be used as a default)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, IEnumerable<Expression> arguments, CallExpressionSegment.ArgumentBracketPresenceOptions? zeroArgBrackets)
		{
			if ((memberAccessTokens.Count() == 1) && !arguments.Any())
			{
				if (memberAccessTokens.Single() is NumericValueToken)
					return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
				if (memberAccessTokens.Single() is DateLiteralToken)
					return new DateValueExpressionSegment(memberAccessTokens.Single() as DateLiteralToken);
				if (memberAccessTokens.Single() is StringToken)
					return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
			}

			CallExpressionSegment.ArgumentBracketPresenceOptions? argBrackets;
			if (arguments.Any())
				argBrackets = null;
			else if (zeroArgBrackets == null)
				argBrackets = CallExpressionSegment.ArgumentBracketPresenceOptions.Absent;
			else
				argBrackets = zeroArgBrackets;

			if (memberAccessTokens.Any())
			{
				return new CallExpressionSegment(
					memberAccessTokens,
					arguments,
					argBrackets
				);
			}
			return new CallSetItemExpressionSegment(
				memberAccessTokens,
				arguments,
				argBrackets
			);
		}

		/// <summary>
		/// Create a CallExpressionSegment from member access tokens and argument expressions (the zeroArgBrackets is only considered if arguments is an empty set,
		/// if arguments is empty and zeroArgBrackets is null then a Absent will be used as a default)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, CallExpressionSegment.ArgumentBracketPresenceOptions? zeroArgBrackets, params Expression[] arguments)
		{
			return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments, zeroArgBrackets);
		}

		/// <summary>
		/// Create a CallExpressionSegment from member access tokens and argument expressions (applying the default logic for ArgumentBracketPresenceOptions; null
		/// if there are arguments and Absent otherwise)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params Expression[] arguments)
		{
			return CALL(memberAccessTokens, (IEnumerable<Expression>)arguments, null);
		}

		private static IExpressionSegment CALLARGSONLY(params IEnumerable<IToken>[] arguments)
		{
			return CALL(new IToken[0], arguments);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token and argument expressions (applying the default logic for ArgumentBracketPresenceOptions;
		/// null if there are arguments and Absent otherwise)
		/// </summary>
		private static IExpressionSegment CALL(IToken memberAccessToken, params Expression[] arguments)
		{
			return CALL(new[] { memberAccessToken }, arguments);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token with no argument expressions and an explicit ArgumentBracketPresenceOptions value
		/// </summary>
		private static IExpressionSegment CALL(IToken memberAccessToken, CallExpressionSegment.ArgumentBracketPresenceOptions zeroArgBrackets)
		{
			return CALL(new[] { memberAccessToken }, new Expression[0], zeroArgBrackets);
		}

		/// <summary>
		/// Create a CallExpressionSegment from a single member access token and argument expressions expressed as token sets (applying the default logic for
		/// ArgumentBracketPresenceOptions; null if there are arguments and Absent otherwise)
		/// </summary>
		private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
		{
			if ((memberAccessTokens.Count() == 1) && !arguments.Any())
			{
				if (memberAccessTokens.Single() is NumericValueToken)
					return new NumericValueExpressionSegment(memberAccessTokens.Single() as NumericValueToken);
				if (memberAccessTokens.Single() is DateLiteralToken)
					return new DateValueExpressionSegment(memberAccessTokens.Single() as DateLiteralToken);
				if (memberAccessTokens.Single() is StringToken)
					return new StringValueExpressionSegment(memberAccessTokens.Single() as StringToken);
			}
			return CALL(
				memberAccessTokens,
				arguments.Select(a => new Expression(new[] { CALL(a) })),
				null
			);
		}

		private static NewInstanceExpressionSegment NEW(string className, int lineIndex)
		{
			return new NewInstanceExpressionSegment(new NameToken(className, lineIndex));
		}

		private static OperationExpressionSegment OP(OperatorToken token)
		{
			return new OperationExpressionSegment(token);
		}

		/// <summary>
		/// Create an Expression from multiple ExpressionSegments
		/// </summary>
		private static Expression EXP(params IExpressionSegment[] segments)
		{
			return new Expression(segments);
		}

		/// <summary>
		/// Create an Expression from a single ExpressionSegment
		/// </summary>
		private static Expression EXP(IExpressionSegment segment)
		{
			return EXP(new[] { segment });
		}
	}
}
