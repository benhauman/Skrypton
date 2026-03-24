namespace Skrypton.CSharpWriter.CodeTranslation.BlockTranslators
{
    internal interface ITranslatorOptions
    {
        bool AcceptTranslationError(string errorKey);
    }
}