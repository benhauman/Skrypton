
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
                new ICodeBlock[]
                {
                    new Statement(
                        new IToken[]
                        {
                            new NameToken("WScript", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NameToken("Echo", 0),
                            new NumericValueToken("1", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NumericValueToken("1", 0)
                        },
                        Statement.CallPrefixOptions.Absent
                    )
                },
                Parser.Parse(TestCulture, "WScript.Echo 1.1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void StatementWithMemberAccessAndDecimalValueAndWrappedMethodArgument()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new Statement(
                        new IToken[]
                        {
                            new NameToken("WScript", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NameToken("Echo", 0),
                            new OpenBrace(0),
                            new NumericValueToken("1", 0),
                            new MemberAccessorOrDecimalPointToken(".", 0),
                            new NumericValueToken("1", 0),
                            new CloseBrace(0)
                        },
                        Statement.CallPrefixOptions.Absent
                    )
                },
                Parser.Parse(TestCulture, "WScript.Echo(1.1)"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void SingleValueSetToNothing()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new[]
                        {
                            new NameToken("a", 0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse(TestCulture, "Set a = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNumber()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new NumericValueToken("1", 0),
                        }),
                        ValueSetTypeOptions.Let
                    )
                },
                Parser.Parse(TestCulture, "a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNumberWithExplicitLet()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new NumericValueToken("1", 0),
                        }),
                        ValueSetTypeOptions.Let
                    )
                },
                Parser.Parse(TestCulture, "Let a(0, 0) = 1"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementSetToNothing()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NumericValueToken("0", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse(TestCulture, "Set a(0, 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }

        [TestMethod, MyFact]
        public void TwoDimensionalArrayElementWithMethodCallIndexSetToNothing()
        {
            myAssert.AreEqual(
                new ICodeBlock[]
                {
                    new ValueSettingStatement(
                        new Expression(new IToken[]
                        {
                            new NameToken("a", 0),
                            new OpenBrace(0),
                            new NameToken("GetValue", 0),
                            new OpenBrace(0),
                            new NumericValueToken("1", 0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("3", 0),
                            new CloseBrace(0),
                            new ArgumentSeparatorToken(0),
                            new NumericValueToken("0", 0),
                            new CloseBrace(0),
                        }),
                        new Expression(new[]
                        {
                            new BuiltInValueToken("Nothing", 0),
                        }),
                        ValueSetTypeOptions.Set
                    )
                },
                Parser.Parse(TestCulture, "Set a(GetValue(1, 3), 0) = Nothing"),
                new CodeBlockSetComparer()
            );
        }
    }
}
