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

        public static string NormalizeUnicodeNarrowNoBreakSpace(this string text)
        {
            /*
                Ubuntu is inserting: U+202F  (NARROW NO-BREAK SPACE)
                    => On Linux (and some newer .NET globalization implementations), date/time formatting follows CLDR / ICU rules instead of the older Windows NLS rules.
                       In CLDR, the AM/PM separator is defined using a narrow no-break space (U+202F) in some cultures.
                Windows: U+0020  (normal ASCII space)
                    => Windows historically used a regular space.
             */
            return text?.Replace('\u202F', ' ');
        }
    }
}