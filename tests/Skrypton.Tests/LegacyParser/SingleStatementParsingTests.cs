
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.LegacyParser;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.Tests.LegacyParser.Helpers;
//#using Xunit#;

namespace Skrypton.Tests.LegacyParser
{
    [TestClass]
    public class SingleStatementParsingTests : TestBase
    {
        [TestMethod, MyFact]
        public void StatementWithMemberAccessAndDecimalValueAndUnwrappedMethodArgument()
        {
            myAssert.AreEqual(
                [
                    new Statement(
                        [
                            new NameToken("WScript", lineIndex1),
                            new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                            new NameToken("Echo", lineIndex1),
                            new NumericValueToken("1", lineIndex1),
                            new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                            new NumericValueToken("1", lineIndex1)
                        ],
                        Statement.CallPrefixOptions.Absent
                    )
                ],
                Parser.Parse(TestCulture, "WScript.Echo 1.1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void StatementWithMemberAccessAndDecimalValueAndWrappedMethodArgument()
        {
            myAssert.AreEqual(
                [
                    new Statement(
                        [
                            new NameToken("WScript", lineIndex1),
                            new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                            new NameToken("Echo", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NumericValueToken("1", lineIndex1),
                            new MemberAccessorOrDecimalPointToken(".", hasLeadingWhiteSpace: false, lineIndex1),
                            new NumericValueToken("1", lineIndex1),
                            new CloseBrace(lineIndex1)
                        ],
                        Statement.CallPrefixOptions.Absent
                    )
                ],
                Parser.Parse(TestCulture, "WScript.Echo(1.1)"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void SingleValueSetToNothing()
        {
            myAssert.AreEqual(
                [
                    new ValueSettingStatement(
                        new CodeExpression([new NameToken("a", lineIndex1)]),
                        new CodeExpression([new BuiltInValueToken("Nothing", lineIndex1)]),
                        ValueSetTypeOptions.Set
                    )
                ],
                Parser.Parse(TestCulture, "Set a = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNumber()
        {
            myAssert.AreEqual(
                [
                    new ValueSettingStatement(
                        new CodeExpression([
                            new NameToken("a", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new ArgumentSeparatorToken(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new CloseBrace(lineIndex1)
                        ]),
                        new CodeExpression([
                            new NumericValueToken("1", lineIndex1)
                        ]),
                        ValueSetTypeOptions.Let
                    )
                ],
                Parser.Parse(TestCulture, "a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNumberWithExplicitLet()
        {
            myAssert.AreEqual(
                [
                    new ValueSettingStatement(
                        new CodeExpression([
                            new NameToken("a", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new ArgumentSeparatorToken(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new CloseBrace(lineIndex1)
                        ]),
                        new CodeExpression([
                            new NumericValueToken("1", lineIndex1)
                        ]),
                        ValueSetTypeOptions.Let
                    )
                ],
                Parser.Parse(TestCulture, "Let a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNothing()
        {
            myAssert.AreEqual(
                [
                    new ValueSettingStatement(
                        new CodeExpression([
                            new NameToken("a", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new ArgumentSeparatorToken(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new CloseBrace(lineIndex1)
                        ]),
                        new CodeExpression([
                            new BuiltInValueToken("Nothing", lineIndex1)
                        ]),
                        ValueSetTypeOptions.Set
                    )
                ],
                Parser.Parse(TestCulture, "Set a(0, 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementWithMethodCallIndexSetToNothing()
        {
            myAssert.AreEqual(
                [
                    new ValueSettingStatement(
                        new CodeExpression([
                            new NameToken("a", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NameToken("GetValue", lineIndex1),
                            new OpenBrace(lineIndex1),
                            new NumericValueToken("1", lineIndex1),
                            new ArgumentSeparatorToken(lineIndex1),
                            new NumericValueToken("3", lineIndex1),
                            new CloseBrace(lineIndex1),
                            new ArgumentSeparatorToken(lineIndex1),
                            new NumericValueToken("0", lineIndex1),
                            new CloseBrace(lineIndex1)
                        ]),
                        new CodeExpression([
                            new BuiltInValueToken("Nothing", lineIndex1)
                        ]),
                        ValueSetTypeOptions.Set
                    )
                ],
                Parser.Parse(TestCulture, "Set a(GetValue(1, 3), 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }
    }
}
