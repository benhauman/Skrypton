using System;
using System.Linq;
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

        protected string BuildTargetNotNullCheckCodeFragment(string targetAccessorName)
        {
            if (targetAccessorName == _envRefName.Name) // _env.
            {
                return string.Empty;
            }
            if (targetAccessorName == _supportRefName.Name) // _.
            {
                return string.Empty;
            }
            if (targetAccessorName == _outerRefName.Name) // _outer.
            {
                return string.Empty;
            }
            string targetNameForException = GetTargetNameForException(targetAccessorName);
            string targetNotNullCheckCs = $@" ?? throw new InvalidOperationException(""Reference not set:{targetNameForException}"")";
            return targetNotNullCheckCs;
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
                throw new NotImplementedException(targetAccessorName);
            }
            else if (targetAccessorName.Contains('"', StringComparison.Ordinal))
            {
                throw new NotImplementedException(targetAccessorName);
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
                throw new NotImplementedException(targetAccessorName);
            }
            return dotTokens[1];
        }
    }
}