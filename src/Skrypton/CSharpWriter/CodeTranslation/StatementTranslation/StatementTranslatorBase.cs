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


        // private readonly CSharpName _supportRefName, _envRefName, _outerRefName;
        protected static string GetTargetNameForException(string targetAccessorName)
        {
            // StatementTranslator
            if (targetAccessorName.Contains('.') || targetAccessorName.Contains('"'))
            {
                return "";
            }
            else
            {
                return targetAccessorName;
            }
        }
    }
}