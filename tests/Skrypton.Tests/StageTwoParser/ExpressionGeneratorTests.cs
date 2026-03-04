
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
    public class ExpressionGeneratorTests : TestBase
    {
        [TestMethod, MyFact]
        public void DirectFunctionCallWithNoArgumentsAndNoBrackets()
        {
            // Test
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("Test", lineIndex1))
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("Test", lineIndex1), CallExpressionSegment.ArgumentBracketPresenceOptions.Present)
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("Test", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("a", lineIndex1), new NameToken("Test", lineIndex1)]
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("a", lineIndex1), new NameToken("b", lineIndex1), new NameToken("Test", lineIndex1)]
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("Test", lineIndex1)],
                            [new NumericValueToken("1", lineIndex1)]
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("Test", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("Test", lineIndex1)],
                            [new NumericValueToken("1", lineIndex1)],
                            [new NumericValueToken("2", lineIndex1)]
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("Test", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new ArgumentSeparatorToken(lineIndex1),
                        new NumericValueToken("2", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("Test", lineIndex1)],
                            EXP(
                                CALL(
                                    [new NameToken("Test2", lineIndex1)],
                                    [new NumericValueToken("1", lineIndex1)]
                                )
                            ),
                            EXP(CALL(new NumericValueToken("2", lineIndex1)))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("Test", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("Test2", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new ArgumentSeparatorToken(lineIndex1),
                        new NumericValueToken("2", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALLSET(
                            CALL(
                                [new NameToken("a", lineIndex1)],
                                [new NumericValueToken("0", lineIndex1)]
                            ),
                            CALL(
                                [new NameToken("Test", lineIndex1)]
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALLSET(
                            CALL(
                                [new NameToken("a", lineIndex1), new NameToken("b", lineIndex1)],
                                [new NumericValueToken("0", lineIndex1)]
                            ),
                            CALL(
                                [new NameToken("Test", lineIndex1)]
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALLSET(
                            CALL(
                                [new NameToken("a", lineIndex1)],
                                [new NumericValueToken("0", lineIndex1)]
                            ),
                            CALL(
                                [new NameToken("b", lineIndex1), new NameToken("Test", lineIndex1)]
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate([
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALLSET(
                            CALL(
                                [new NameToken("a", lineIndex1)],
                                [new NumericValueToken("0", lineIndex1)]
                            ),
                            CALLARGSONLY(
                                [new NumericValueToken("1", lineIndex1)]
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1),
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        BR(
                            CALL(new NameToken("a", lineIndex1)),
                            OP(new OperatorToken("+", lineIndex1)),
                            CALL(new NameToken("b", lineIndex1))
                        ),
                        OP(new OperatorToken("+", lineIndex1)),
                        CALL(new NameToken("c", lineIndex1))
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("c", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new OperatorToken("+", lineIndex1)),
                        BR(
                            CALL(new NameToken("b", lineIndex1)),
                            OP(new OperatorToken("*", lineIndex1)),
                            CALL(new NameToken("c", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new NameToken("c", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new OperatorToken("+", lineIndex1)),
                        BR(
                            CALL(new NameToken("b", lineIndex1)),
                            OP(new OperatorToken("*", lineIndex1)),
                            CALL(
                                [new NameToken("c", lineIndex1)],
                                [new NumericValueToken("0", lineIndex1)]
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new NameToken("c", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        BR(
                            CALL(new NameToken("a", lineIndex1)),
                            OP(new OperatorToken("+", lineIndex1)),
                            BR(
                                CALL(new NameToken("b", lineIndex1)),
                                OP(new OperatorToken("*", lineIndex1)),
                                CALL(
                                    [new NameToken("c", lineIndex1)],
                                    [new NumericValueToken("0", lineIndex1)]
                                )
                            )
                        ),
                        OP(new OperatorToken("+", lineIndex1)),
                        CALL(new NameToken("d", lineIndex1))
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new NameToken("c", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("d", lineIndex1),
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new OperatorToken("+", lineIndex1)),
                        BR(
                            CALL(new NameToken("b", lineIndex1)),
                            OP(new OperatorToken("*", lineIndex1)),
                            CALL(new NameToken("c", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new NameToken("c", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        BR(
                            CALL(new NameToken("a", lineIndex1)),
                            OP(new OperatorToken("+", lineIndex1)),
                            BR(
                                CALL(new NameToken("b", lineIndex1)),
                                OP(new OperatorToken("*", lineIndex1)),
                                CALL(new NameToken("c", lineIndex1))
                            )
                        ),
                        OP(new OperatorToken("+", lineIndex1)),
                        CALL(new NameToken("d", lineIndex1))
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new NameToken("c", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("d", lineIndex1),
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new ComparisonOperatorToken("=", lineIndex1)),
                        BR(
                            CALL(new NameToken("b", lineIndex1)),
                            OP(new OperatorToken("+", lineIndex1)),
                            CALL(new NameToken("c", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new ComparisonOperatorToken("=", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("c", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        BR(
                            BR(
                                CALL(new NameToken("a", lineIndex1)),
                                OP(new OperatorToken("+", lineIndex1)),
                                BR(
                                    CALL(new NameToken("b", lineIndex1)),
                                    OP(new OperatorToken("*", lineIndex1)),
                                    CALL(
                                        [new NameToken("c", lineIndex1), new NameToken("d", lineIndex1)],
                                        EXP(
                                            CALL(
                                                [new NameToken("Test", lineIndex1)],
                                                [new NumericValueToken("0", lineIndex1)]
                                            )
                                        ),
                                        EXP(CALL(new NumericValueToken("1", lineIndex1)))
                                    )
                                )
                            ),
                            OP(new OperatorToken("+", lineIndex1)),
                            CALL(new NameToken("e", lineIndex1))
                        ),
                        OP(new ComparisonOperatorToken("=", lineIndex1)),
                        CALL(new NameToken("f", lineIndex1))
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("b", lineIndex1),
                        new OperatorToken("*", 1),
                        new NameToken("c", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("d", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("Test", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("0", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new ArgumentSeparatorToken(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new OperatorToken("+", lineIndex1),
                        new NameToken("e", lineIndex1),
                        new ComparisonOperatorToken("=", lineIndex1),
                        new NameToken("f", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new OperatorToken("*", lineIndex1)),
                        BR(
                            OP(new OperatorToken("-", lineIndex1)),
                            CALL(new NameToken("b", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new OperatorToken("*", lineIndex1),
                        new OperatorToken("-", lineIndex1),
                        new NameToken("b", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new LogicalOperatorToken("AND", lineIndex1)),
                        BR(
                            OP(new LogicalOperatorToken("NOT", lineIndex1)),
                            CALL(new NameToken("b", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new LogicalOperatorToken("AND", lineIndex1),
                        new LogicalOperatorToken("NOT", lineIndex1),
                        new NameToken("b", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(new NameToken("a", lineIndex1)),
                        OP(new LogicalOperatorToken("AND", lineIndex1)),
                        BR(
                            OP(new LogicalOperatorToken("NOT", lineIndex1)),
                            BR(
                                OP(new LogicalOperatorToken("NOT", lineIndex1)),
                                CALL(new NameToken("b", lineIndex1))
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("a", lineIndex1),
                        new LogicalOperatorToken("AND", lineIndex1),
                        new LogicalOperatorToken("NOT", lineIndex1),
                        new LogicalOperatorToken("NOT", lineIndex1),
                        new NameToken("b", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        OP(new LogicalOperatorToken("NOT", lineIndex1)),
                        BR(
                            CALL(new NameToken("a", lineIndex1)),
                            OP(new ComparisonOperatorToken("IS", lineIndex1)),
                            new BuiltInValueExpressionSegment(new BuiltInValueToken("Nothing", lineIndex1))
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new LogicalOperatorToken("NOT", lineIndex1),
                        new NameToken("a", lineIndex1),
                        new ComparisonOperatorToken("IS", lineIndex1),
                        new BuiltInValueToken("Nothing", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        NEW("Test", lineIndex1)
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new KeyWordToken("new", lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            new NameToken("Test", lineIndex1),
                            EXP(
                                BR(CALL(new NameToken("a", lineIndex1)))
                            )
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                        [
                            new NameToken("Test", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new OpenBrace(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new DoNotRenameNameToken("a", lineIndex1), new NameToken("Test", lineIndex1)]
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("Test", lineIndex1)
                    ],
                    directedWithReferenceIfAny: new DoNotRenameNameToken("a", lineIndex1),
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
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorToken(0),
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
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
                [new NumericValueToken("1", lineIndex1), new OpenBrace(lineIndex1), new CloseBrace(lineIndex1)],
                typeof(TypeMismatchException),
                "'[number: 1]' is called like a function"
            );
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("wscript", lineIndex1), new NameToken("echo", lineIndex1)],
                            new Expression([runtimeErrorExpressionSegment])
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
                    directedWithReferenceIfAny: new DoNotRenameNameToken("a", lineIndex1),
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
                     [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                     ],
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
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("wscript", lineIndex1), new NameToken("echo", lineIndex1)],
                            new Expression([
                                CALL([new StringToken("1", lineIndex1), new NameToken("a", lineIndex1)])
                            ])
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new StringToken("1", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
                    directedWithReferenceIfAny: new DoNotRenameNameToken("a", lineIndex1),
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
                [new NumericValueToken("1", lineIndex1), new OpenBrace(lineIndex1), new CloseBrace(lineIndex1)],
                typeof(TypeMismatchException),
                "'[number: 1]' is called like a function"
            );
            myAssert.AreEqual(
                [
                    EXP(
                        CALL(
                            [new NameToken("wscript", lineIndex1), new NameToken("echo", lineIndex1)],
                            new Expression([runtimeErrorExpressionSegment])
                        )
                    )
                ],
                ExpressionGenerator.Generate(
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
                    directedWithReferenceIfAny: new DoNotRenameNameToken("a", lineIndex1),
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
                    [
                        new NameToken("wscript", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("echo", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NumericValueToken("1", lineIndex1),
                        new MemberAccessorToken(lineIndex1),
                        new NameToken("a", lineIndex1),
                        new OpenBrace(lineIndex1),
                        new NameToken("b", lineIndex1),
                        new CloseBrace(lineIndex1),
                        new CloseBrace(lineIndex1)
                    ],
                    directedWithReferenceIfAny: null,
                    warningLogger: warning => { }
                );
            });
        }

        // TODO: Built-in constants and boolean member access attempts (these are consistent with string literals in all cases)

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        //private static BracketedExpressionSegment BR(IReadOnlyCollection<IExpressionSegment> segments)
        //{
        //	return new BracketedExpressionSegment(segments);
        //}

        /// <summary>
        /// Create a BracketedExpressionSegment from a set of expressions
        /// </summary>
        private static BracketedExpressionSegment BR(params IExpressionSegment[] segments)
        {
            return new BracketedExpressionSegment((IReadOnlyCollection<IExpressionSegment>)segments);
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
            return CALL([], arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions (applying the default logic for ArgumentBracketPresenceOptions;
        /// null if there are arguments and Absent otherwise)
        /// </summary>
        private static IExpressionSegment CALL(IToken memberAccessToken, params Expression[] arguments)
        {
            return CALL([memberAccessToken], arguments);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token with no argument expressions and an explicit ArgumentBracketPresenceOptions value
        /// </summary>
        private static IExpressionSegment CALL(IToken memberAccessToken, CallExpressionSegment.ArgumentBracketPresenceOptions zeroArgBrackets)
        {
            return CALL([memberAccessToken], [], zeroArgBrackets);
        }

        /// <summary>
        /// Create a CallExpressionSegment from a single member access token and argument expressions expressed as token sets (applying the default logic for
        /// ArgumentBracketPresenceOptions; null if there are arguments and Absent otherwise)
        /// </summary>
        private static IExpressionSegment CALL(IEnumerable<IToken> memberAccessTokens, params IEnumerable<IToken>[] arguments)
        {
            if ((memberAccessTokens.Count() == 1) && arguments.Length == 0)
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
                arguments.Select(a => new Expression([CALL(a)])),
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
            return EXP([segment]);
        }
    }
}
