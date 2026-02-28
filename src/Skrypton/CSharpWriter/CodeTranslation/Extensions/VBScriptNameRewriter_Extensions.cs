using System;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.RuntimeSupport;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    public static class VBScriptNameRewriter_Extensions
    {
        /// <summary>
        /// When trying to access variables, functions, classes, etc.. we need to pass the member's name through the VBScriptNameRewriter. In
        /// most cases this token will be a NameToken which we can pass straight in, but in some cases it may be another type (perhaps a key
        /// word type) and so will have to be wrapped in a NameToken instance before passing through the name rewriter. This extension
		/// method should be used in all places where the VBScriptNameRewriter is used by the CSharpWriter since it allows us to override
		/// its behaviour where required - eg. by using a DoNotRenameNameToken
        /// </summary>
        public static string GetMemberAccessTokenName(this VBScriptNameRewriter nameRewriter, IToken token)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException(nameof(nameRewriter));
            if (token == null)
                throw new ArgumentNullException(nameof(token));

            // A TargetCurrentClassToken indicates a "Me" (eg. "Me.Name") which can always be translated directly into "this". In VBScript,
            // "Me" is valid even when not explicitly within a VBScript class (it refers to the outermost scope, so "Me.F1()" will try to
            // call a function "F1" in the outermost scope). When the code IS explicitly within a VBScript class, the "Me" reference is
            // the instance of that class. In the translated code, both cases are fine to translate straight into whatever "this" is
            // at runtime.
            if (token is TargetCurrentClassToken)
                return "this";

            var nameToken = (token as NameToken) ?? new ForRenamingNameToken(token.ContentUpperX(), token.LineIndex);
            if (nameToken is DoNotRenameNameToken)
                return nameToken.Content;
            if (token is BuiltInFunctionToken bfun)
            {
                if (DefaultRuntimeSupportClassFactory._caseInsensitiveCSharpKeywordMatcher.Contains(nameToken.Content))
                {
                    //if (string.Equals("int", nameToken.Content, StringComparison.OrdinalIgnoreCase))
                    //    return $"int__on_line_{token.LineIndex}"; // dirty fix for : Int ii = 0
                    // ?!? Int => "rewritten_int"
                    if (string.Equals("Int", nameToken.Content, StringComparison.Ordinal))
                    {
                        // test with 'UnitSelection_Renderer_NoSelects'
                        // Int(...) in VBScript is a numeric function that takes a number and returns the largest whole number less than or equal to it — in other words, it truncates toward negative infinity.
                    }
                    else if (string.Equals("Join", nameToken.Content, StringComparison.Ordinal))
                    {
                        // test with 'UnitSelection_Renderer_NoSelects'
                        // Join(aryFormattedData, "") is a VBScript string‑building function.
                        // It takes an array of strings and concatenates them into one big string, using the second argument as the separator.
                        // example: Join(Array("a","b","c"), ",") => a,b,c
                    }
                    else
                    {
                        // VBScript built‑in functions that appear in your list:
                        // 'Join' → VBScript string function
                        // 'Int' → VBScript numeric function(C# has int keyword)
                        // 'String' → VBScript function String(n, char)(C# keyword string)
                        // 'TypeOf' → VBScript operator TypeOf x Is Something(C# keyword typeof)
                        // 'Is' → VBScript comparison operator (C# pattern matching operator)

                        throw new InvalidOperationException($"Invalid name or keyword: '{nameToken.Content}'. Line:{token.LineIndex}");
                    }
                }
            }
            return nameRewriter.RewriteVBScriptName(nameToken).Name;
        }

        public static bool AreNamesEquivalent(this VBScriptNameRewriter nameRewriter, NameToken x, NameToken y)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException(nameof(nameRewriter));
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            return nameRewriter.GetMemberAccessTokenName(x) == nameRewriter.GetMemberAccessTokenName(y);
        }

        /// <summary>
        /// This is used by the GetMemberAccessTokenName for tokens that are not already NameToken instances. This derived type is used
        /// since it will bypass some of the the validation in the NameToken base constructor.
        /// </summary>
        private sealed class ForRenamingNameToken : NameToken
        {
            public ForRenamingNameToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex) { }
        }
    }
}
