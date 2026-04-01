using System;
using System.Linq;

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


        protected internal string GetTargetNameForException(string targetAccessorName)
        {
            if (targetAccessorName == null) throw new ArgumentNullException(nameof(targetAccessorName));
            // StatementTranslator
            string[] dotTokens = targetAccessorName.Split('.');
            if (dotTokens.Length > 1)
            {
                if (dotTokens.Length == 2 && dotTokens[0] == _envRefName.Name) // _env.
                {
                    if (dotTokens[1].Contains('"'))
                    {
                        return "";
                    }
                    return dotTokens[1];
                }
                if (dotTokens.Length == 2 && dotTokens[0] == _supportRefName.Name) // _.
                {
                    if (dotTokens[1].Contains('"'))
                    {
                        return "";
                    }
                    return dotTokens[1];
                }
                if (dotTokens.Length == 2 && dotTokens[0] == _outerRefName.Name) // _outer.
                {
                    if (dotTokens[1].Contains('"'))
                    {
                        return "";
                    }
                    return dotTokens[1];
                }
                if (targetAccessorName.StartsWith($"{_supportRefName.Name}.CALL", StringComparison.Ordinal)) // _.CALLm1v1(this, _outer.rs ?? throw new InvalidOperationException("Reference not set:rs"), "fields", (Int16)0)
                {
                    return "(call result)";
                }
                throw new NotImplementedException(targetAccessorName);
                //return "";
            }
            else if (targetAccessorName.Contains('"'))
            {
                throw new NotImplementedException(targetAccessorName);
                //return "";
            }
            else
            {
                return targetAccessorName;
            }
        }
    }
}