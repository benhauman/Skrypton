using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass]
    public sealed class MyVBScriptRegExpTests : TestBase
    {
        // -------------------------------------------------------------
        //  BASIC FUNCTIONALITY
        // -------------------------------------------------------------
        private MyVBScriptRegExp NewMyVBScriptRegExp() => new MyVBScriptRegExp(base.TestCulture);

        [TestMethod]
        public void Test_SimpleMatch()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "abc";

            Assert.IsTrue(re.Test("123abc456"));
            Assert.IsFalse(re.Test("zzz"));
        }

        [TestMethod]
        public void Test_IgnoreCase()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "abc";

            re.IgnoreCase = false;
            Assert.IsFalse(re.Test("ABC"));

            re.IgnoreCase = true;
            Assert.IsTrue(re.Test("ABC"));
        }

        // -------------------------------------------------------------
        //  GLOBAL VS NON-GLOBAL
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_Replace_NonGlobal()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "\\d+";
            re.Global = false;

            var result = re.Replace("a1b2c3", "#");
            Assert.AreEqual("a#b2c3", result);  // only first replaced
        }

        [TestMethod]
        public void Test_Replace_Global()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "\\d+";
            re.Global = true;

            var result = re.Replace("a1b2c3", "#");
            Assert.AreEqual("a#b#c#", result);  // all replaced
        }

        // -------------------------------------------------------------
        //  EXECUTE() AND MATCH COLLECTION
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_Execute_ReturnsCorrectMatches()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "\\d+";

            var matches = re.Execute("x1y22z333");

            Assert.AreEqual(3, matches.Count);

            Assert.AreEqual("1", matches[0].Value);
            Assert.AreEqual("22", matches[1].Value);
            Assert.AreEqual("333", matches[2].Value);

            Assert.AreEqual(1, matches[0].Length);
            Assert.AreEqual(2, matches[1].Length);
            Assert.AreEqual(3, matches[2].Length);
        }

        // -------------------------------------------------------------
        //  SUBMATCHES
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_SubMatches_CaptureGroups()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "(a)(b)(c)";
            re.Global = true;

            var matches = re.Execute("abc abc");

            Assert.AreEqual(2, matches.Count);

            var first = matches[0].SubMatches;
            Assert.AreEqual(3, first.Count);
            Assert.AreEqual("a", first[0]);
            Assert.AreEqual("b", first[1]);
            Assert.AreEqual("c", first[2]);
        }

        // -------------------------------------------------------------
        //  MULTILINE
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_MultilineBehavior()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "^abc";   // anchor
            re.Multiline = 0;

            Assert.IsFalse(re.Test("zzz\nabc"));
            Assert.IsTrue(re.Test("abc\nzzz"));

            re.Multiline = 1;

            Assert.IsTrue(re.Test("zzz\nabc"));
        }

        // -------------------------------------------------------------
        //  COMPARE MODE (Binary vs Text)
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_Compare_TextMode_IgnoresCulture()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "straße";   // German sharp-s

            // Binary compare: case-sensitive literal
            re.Compare = 0;
            Assert.IsTrue(re.Test("straße"));
            Assert.IsFalse(re.Test("STRASSE"));

            // Text compare: culture-invariant + case-insensitive-ish behavior
            re.IgnoreCase = true;
            //re.Compare = 1;
            Assert.IsTrue(re.Test("STRASSE"));
        }

        // -------------------------------------------------------------
        //  REPLACE LITERAL EDGE CASES
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_Replace_EmptyReplacement()
        {
            // In VBScript RegExp, Replace only replaces the first match unless you set Global = true. That’s why you got "ab2" instead of "ab" — only the 1 was removed, the 2 remained.

            var re = NewMyVBScriptRegExp();
            re.Pattern = "\\d+";
            re.Global = true;          // <-- Required to replace all matches

            var result = re.Replace("a1b2", "");

            Assert.AreEqual("ab", result);
        }

        // -------------------------------------------------------------
        //  EMPTY INPUT AND NULL SAFETY
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_NullInput()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "abc";

            Assert.IsFalse(re.Test(null));

            var result = re.Replace(null, "X");
            Assert.AreEqual("", result);
        }

        // -------------------------------------------------------------
        //  REAL-WORLD VBScript EXAMPLES
        // -------------------------------------------------------------

        [TestMethod]
        public void Test_RealVBScriptExample_MatchesWords()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "\\w+";
            re.Global = true;

            var matches = re.Execute("hello world 123");

            Assert.AreEqual(3, matches.Count);
            Assert.AreEqual("hello", matches[0].Value);
            Assert.AreEqual("world", matches[1].Value);
            Assert.AreEqual("123", matches[2].Value);
        }

        [TestMethod]
        public void Test_RealVBScriptExample_ReplaceEmail()
        {
            var re = NewMyVBScriptRegExp();
            re.Pattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+";
            re.Global = true;

            var result = re.Replace("Email me at test@example.com please", "[hidden]");

            Assert.AreEqual("Email me at [hidden] please", result);
        }
    }
}
