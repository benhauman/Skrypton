using System;
using System.Linq;
using Skrypton.CSharpWriter.CodeTranslation.Extensions;
using Skrypton.RuntimeSupport;

namespace Skrypton.CSharpWriter.CodeTranslation.StatementTranslation
{
    public abstract class StatementTranslatorBase
    {
#pragma warning disable CA1051 // Do not declare visible instance fields
        protected internal readonly CSharpName _supportRefName, _envRefName, _outerRefName;
#pragma warning restore CA1051 // Do not declare visible instance fields

        protected StatementTranslatorBase(
            CSharpName supportRefName,
            CSharpName envRefName,
            CSharpName outerRefName
            )
        {
            _supportRefName = supportRefName ?? throw new ArgumentNullException(nameof(supportRefName));
            _envRefName = envRefName ?? throw new ArgumentNullException(nameof(envRefName));
            _outerRefName = outerRefName ?? throw new ArgumentNullException(nameof(outerRefName));
        }

        internal string RenderNnO(string targetAccessorName, string? targetNameForException = null)
        {
            if (targetAccessorName == _envRefName.Name) // _env.
            {
                return targetAccessorName;
            }
            if (targetAccessorName == _supportRefName.Name) // _.
            {
                //throw new NotImplementedException();// do not render it via "CALL/SET" but directly like: _.RAISE
                return targetAccessorName;
            }
            if (targetAccessorName == _outerRefName.Name) // _outer.
            {
                return targetAccessorName;
            }

            // Error 424 names a local variable or parameter, but for a global variable (or the result of a support function) VBScript names its
            // value - so only locals pass their name (issue #84)
            string firstToken = targetAccessorName.Split('.')[0];
            bool isGlobal = (targetNameForException == null) && ((firstToken == _envRefName.Name) || (firstToken == _outerRefName.Name) || (firstToken == _supportRefName.Name));
            string targetNameLiteral = isGlobal ? "null" : (targetNameForException ?? GetTargetNameForException(targetAccessorName)).ToLiteral();
            return $"_.{nameof(IProvideVBScriptCompatFunctionalityToIndividualRequests.NnO)}({targetAccessorName}, {targetNameLiteral})";
        }
        private string GetTargetNameForException(string targetAccessorName)
        {
            if (targetAccessorName == null) throw new ArgumentNullException(nameof(targetAccessorName));

            string[] dotTokens = targetAccessorName.Split('.');
            if (dotTokens.Length > 1)
            {
                if (dotTokens.Length >= 2 && dotTokens[0] == _envRefName.Name) // _env.
                {
                    return GetTargetChainTokenAsText(dotTokens, targetAccessorName, "e");
                }
                if (dotTokens.Length >= 2 && dotTokens[0] == _supportRefName.Name) // _.
                {
                    return GetTargetChainTokenAsText(dotTokens, targetAccessorName, "_");
                }
                if (dotTokens.Length >= 2 && dotTokens[0] == _outerRefName.Name) // _outer.
                {
                    return GetTargetChainTokenAsText(dotTokens, targetAccessorName, "o");
                }
                return GetCallResultName(targetAccessorName);
            }
            else if (dotTokens.Length == 1)
            {
                if (targetAccessorName.StartsWith(",", StringComparison.Ordinal))
                {
                    return "(member)";
                }
                if (targetAccessorName.Contains('"', StringComparison.Ordinal))
                {
                    return GetCallResultName(targetAccessorName); // see issue #72
                }
                else
                {
                    return targetAccessorName;
                }
            }
            else if (targetAccessorName.Contains('"', StringComparison.Ordinal))
            {
                return GetCallResultName(targetAccessorName); // see issue #72
            }
            else
            {
                return targetAccessorName;
            }
        }

        private string GetTargetChainTokenAsText(string[] dotTokens, string targetAccessorName, string alias0)
        {
            string token0 = dotTokens[0];
            if (dotTokens[1].Contains('"', StringComparison.Ordinal) || targetAccessorName.Contains('"', StringComparison.Ordinal))
            {
                // test: InvalidFunctionSettingMustCompileThoughFailAtRunTime
                if (targetAccessorName.StartsWith($"{token0}.{nameof(IProvideVBScriptCompatFunctionalityToIndividualRequests.RAISEERROR)}(", StringComparison.Ordinal)) // _.RAISEERROR(new Illegal
                {
                    return "(error result)";
                }
                if (targetAccessorName.StartsWith($"{token0}.CALL", StringComparison.Ordinal)) // _.CALLm1v1(this, oCaseType
                {
                    return $"({alias0}.call result)";
                }
                if (targetAccessorName.StartsWith($"{token0}.{nameof(IProvideVBScriptCompatFunctionalityToIndividualRequests.GETOBJECT)}", StringComparison.Ordinal)) // _.CALLm1v1(this, oCaseType
                {
                    return "(GetObject result)";
                }
                if (targetAccessorName.StartsWith($"{token0}.{nameof(IProvideVBScriptCompatFunctionalityToIndividualRequests.CREATEOBJECT)}", StringComparison.Ordinal)) // _.CREATEOBJECT("..").Member (see issue #72)
                {
                    return "(CreateObject result)";
                }
                return GetCallResultName(targetAccessorName);
            }
            return dotTokens[1];
        }

        /// <summary>
        /// This name is only used in the "Object required" error message for a member access on the target, so any other call result gets a generic
        /// name rather than failing the translation (see issue #72), eg. "_.LCASE(x)" is "(LCASE result)"
        /// </summary>
        private static string GetCallResultName(string targetAccessorName)
        {
            int openBracketIndex = targetAccessorName.IndexOf('(', StringComparison.Ordinal);
            if (openBracketIndex > 0)
            {
                string callee = targetAccessorName.Substring(0, openBracketIndex);
                int lastDotIndex = callee.LastIndexOf('.');
                string name = (lastDotIndex >= 0) ? callee.Substring(lastDotIndex + 1) : callee;
                if (name.Length != 0)
                    return $"({name} result)";
            }
            return "(expression result)";
        }
    }
}