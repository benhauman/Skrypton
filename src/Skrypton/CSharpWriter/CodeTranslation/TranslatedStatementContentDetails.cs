using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("({VariablesAccessed.Count}){TranslatedContent}")]
    public class TranslatedStatementContentDetails // base class of 'TranslatedStatementContentDetailsWithContentType'
    {
        public TranslatedStatementContentDetails(TranslatedStatementContentDetailsKind kind, string translatedContent, IReadOnlyCollection<NameToken> variablesAccessed)
        {
            if (string.IsNullOrWhiteSpace(translatedContent))
                throw new ArgumentException("Null/blank translatedContent specified");
            TranslatedContent = translatedContent;
            VariablesAccessed = variablesAccessed ?? throw new ArgumentNullException(nameof(variablesAccessed));
        }

        /// <summary>
        /// This will never return null or blank
        /// </summary>
        public string TranslatedContent { get; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public IReadOnlyCollection<NameToken> VariablesAccessed { get; private set; }
    }

    public enum TranslatedStatementContentDetailsKind
    {
        Unknown,
        RawText,
        ReturnType,
        Args,
        CallContent,
        CallText,
        ConstNothing,
        ConstTrue,
        ConstFalse,
        ConstKnown,
        DotEqual,
        DotERR,
        DotIf,
        DotNewRefExp,
        DotNewClassInstance,
        DotNullableDate,
        DotNullableNum,
        DotNullableSTR,
        DotNum,
        DotVal,
        DorRaiseError,
        DotRef,
        DotRefIfArray,
        NotUntil,
        IfResultName,
        TargetName,
        ValueDateFromParse,
        ValueNum,
        ValueString,
    }
}
