namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    internal interface ITranslatorOptions
    {
        bool AcceptTranslationError(string errorKey);
        void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex);
    }
}