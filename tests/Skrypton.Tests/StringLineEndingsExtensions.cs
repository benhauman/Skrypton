namespace Skrypton.Tests
{
    internal static class StringLineEndingsExtensions
    {
        public static string NormalizeLineEndings(this string text)
        {
            return text?.Replace("\r\n", "\n").Replace("\r", "\n");
        }

        public static string[] SplitLines(this string text)
        {
            return text.Split(["\r\n", "\n", "\r"], System.StringSplitOptions.None);
        }
        public static string[] SplitLinesRemoveEmptyEntries(this string text)
        {
            return text.Split(["\r\n", "\n", "\r"], System.StringSplitOptions.RemoveEmptyEntries);
        }
    }
}