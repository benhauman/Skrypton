namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    internal interface ITranslatorOptions
    {
        bool AcceptTranslationError(string errorKey);
        void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex);

        /// <summary>
        /// The script has "Option Explicit" - undeclared variables are errors (VBScript: "Variable is undefined")
        /// </summary>
        void OptionExplicitDetected();

        /// <summary>
        /// VBScript's rule for undeclared variables is used (see DefaultTranslatorOptions) - an undeclared variable with Option Explicit is then
        /// VBScript's error 500 "Variable is undefined"
        /// </summary>
        bool UsesVBScriptRuleForUndeclaredVariables { get; }
    }
}
