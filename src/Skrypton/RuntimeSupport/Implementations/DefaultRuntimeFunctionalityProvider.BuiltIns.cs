using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport.Implementations
{
    /// <summary>
    /// VBScript built-in functions that were missing (see issue #79) - their behaviour is that of Windows VBScript (ComParity_BuiltInFunctions.vbs)
    /// </summary>
    public sealed partial class DefaultRuntimeFunctionalityProvider
    {
        private const int VbUseDefault = -2;

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Logical operators

        /// <summary>
        /// l Eqv r - Null if either side is Null, otherwise Not (l Xor r) (Boolean for Booleans, otherwise a number)
        /// </summary>
        public object EQV(object l, object r)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> values = GetForBitwiseOperations("'Eqv'", l, r);
            int? left = values.Item1.First();
            int? right = values.Item1.Skip(1).Single();
            if ((left == null) || (right == null))
                return DBNull.Value;
            return values.Item2(~(left.Value ^ right.Value));
        }

        /// <summary>
        /// l Imp r - (Not l) Or r; with Null: Null Imp True and False Imp Null are True, the other combinations with Null are Null
        /// </summary>
        public object IMP(object l, object r)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> values = GetForBitwiseOperations("'Imp'", l, r);
            int? left = values.Item1.First();
            int? right = values.Item1.Skip(1).Single();
            // with Null the result has the type of the other side (False Imp Null is the Boolean True)
            if (left == null)
                return (right == -1) ? ((_valueRetriever.VAL(r) is bool) ? (object)true : values.Item2(-1)) : DBNull.Value;
            if (right == null)
                return (left == 0) ? ((_valueRetriever.VAL(l) is bool) ? (object)true : values.Item2(-1)) : DBNull.Value;
            return values.Item2(~left.Value | right.Value);
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Numbers

        public object EXP(object value) => Math.Exp(CDBL_Precise(value, "'Exp'"));

        public object LOG(object value)
        {
            double number = CDBL_Precise(value, "'Log'");
            if (number <= 0)
                throw new InvalidProcedureCallOrArgumentException("'Log' (the number must be greater than zero)");
            return Math.Log(number);
        }

        public object TAN(object value) => Math.Tan(CDBL_Precise(value, "'Tan'"));

        /// <summary>
        /// Fix - the integer part (towards zero), of the type of the value (Integer, Long, Double, ..); Null for Null
        /// </summary>
        public object FIX(object value)
        {
            object? number = _valueRetriever.VAL(value, "'Fix'");
            if (number == DBNull.Value)
                return DBNull.Value;
            if (number == null)
                return (short)0;
            if ((number is byte) || (number is short) || (number is int))
                return number;
            if (number is bool boolean)
                return (short)(boolean ? -1 : 0);
            if (number is DateTime date)
                return date.Date;
            if (number is decimal currency)
                return Math.Truncate(currency);
            if (number is float single)
                return (float)Math.Truncate(single);
            return Math.Truncate(GetAsNumber<double>(number, "'Fix'", Convert.ToDouble));
        }

        /// <summary>
        /// Sgn - -1, 0 or 1 (an Integer); Null is "Invalid use of Null"
        /// </summary>
        public object SGN(object value)
        {
            object? number = _valueRetriever.VAL(value, "'Sgn'");
            if (number == DBNull.Value)
                throw new InvalidUseOfNullException("'Sgn'");
            if (number == null)
                return (short)0;
            return (short)Math.Sign(CDBL_Precise(number, "'Sgn'"));
        }

        /// <summary>
        /// Oct - as Hex, in base 8
        /// </summary>
        public object OCT(object value)
        {
            object? number = _valueRetriever.VAL(value, "'Oct'");
            if (number == DBNull.Value)
                return DBNull.Value;
            bool shortForNegativeValues = (number is bool) || (number is short);
            long numeric = CLNG(number, "'Oct'");
            if (numeric < 0)
                numeric += shortForNegativeValues ? 0x10000L : 0x100000000L;
            return Convert.ToString(numeric, 8);
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Strings

        /// <summary>
        /// ChrW - the character with the Unicode code (-32768 to 65535)
        /// </summary>
        public object CHRW(object value)
        {
            object? code = _valueRetriever.VAL(value, "'ChrW'");
            if (code == DBNull.Value)
                throw new InvalidUseOfNullException("'ChrW'");
            int number = CLNG(code, "'ChrW'");
            if ((number < -32768) || (number > 65535))
                throw new InvalidProcedureCallOrArgumentException($"'ChrW' ({number} is not a character code)");
            return new string((char)(number & 0xFFFF), 1);
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Byte strings (issue #93): a string is its UTF-16LE bytes (it may have an odd number of them, see VBScriptByteStrings), a byte array (eg.
        // ADODB.Stream.Read, ServerXMLHTTP.responseBody) is its bytes

        /// <summary>
        /// The bytes of the value - null if it is Null
        /// </summary>
        private byte[]? GetBytesOrNull(object? value, string functionName)
        {
            value = _valueRetriever.VAL(value, functionName);
            if (value == DBNull.Value)
                return null;
            if (value is byte[] bytes)
                return bytes;
            return VBScriptByteStrings.GetBytes(_valueRetriever.STR(value, functionName));
        }

        /// <summary>
        /// LenB - the number of bytes (Null for Null)
        /// </summary>
        public object LENB(object? value)
        {
            byte[]? bytes = GetBytesOrNull(value, "'LenB'");
            return (bytes == null) ? DBNull.Value : (object)bytes.Length;
        }

        /// <summary>
        /// AscB - the first byte (a Byte)
        /// </summary>
        public object ASCB(object value)
        {
            byte[] bytes = GetBytesOrNull(value, "'AscB'") ?? throw new InvalidUseOfNullException("'AscB'");
            if (bytes.Length == 0)
                throw new InvalidProcedureCallOrArgumentException("'AscB' (empty string)");
            return bytes[0];
        }

        /// <summary>
        /// ChrB - a string of one byte
        /// </summary>
        public object CHRB(object value)
        {
            object? code = _valueRetriever.VAL(value, "'ChrB'");
            if (code == DBNull.Value)
                throw new InvalidUseOfNullException("'ChrB'");
            int number = CLNG(code, "'ChrB'");
            if ((number < 0) || (number > 255))
                throw new VBScriptOverflowException($"'ChrB' ({number} is not a byte)");
            return VBScriptByteStrings.GetString(new[] { (byte)number });
        }

        /// <summary>
        /// LeftB - the first bytes
        /// </summary>
        public object LEFTB(object value, object maxLength)
        {
            byte[]? bytes = GetBytesOrNull(value, "'LeftB'");
            int length = CLNG(maxLength, "'LeftB'");
            if (length < 0)
                throw new InvalidProcedureCallOrArgumentException("'LeftB' (length may not be negative)");
            return (bytes == null) ? DBNull.Value : (object)VBScriptByteStrings.GetString(bytes, 0, Math.Min(length, bytes.Length));
        }

        /// <summary>
        /// RightB - the last bytes
        /// </summary>
        public object RIGHTB(object value, object maxLength)
        {
            byte[]? bytes = GetBytesOrNull(value, "'RightB'");
            int length = CLNG(maxLength, "'RightB'");
            if (length < 0)
                throw new InvalidProcedureCallOrArgumentException("'RightB' (length may not be negative)");
            if (bytes == null)
                return DBNull.Value;
            length = Math.Min(length, bytes.Length);
            return VBScriptByteStrings.GetString(bytes, bytes.Length - length, length);
        }

        public object MIDB(object value, object start) => MIDB(value, start, int.MaxValue);

        /// <summary>
        /// MidB - the bytes from the (1-based) start, at most length of them
        /// </summary>
        public object MIDB(object value, object start, object length)
        {
            byte[]? bytes = GetBytesOrNull(value, "'MidB'");
            int startNumber = CLNG(start, "'MidB'");
            int lengthNumber = CLNG(length, "'MidB'");
            if (startNumber <= 0)
                throw new InvalidProcedureCallOrArgumentException("'MidB' (start must be a positive number)");
            if (lengthNumber < 0)
                throw new InvalidProcedureCallOrArgumentException("'MidB' (length may not be negative)");
            if (bytes == null)
                return DBNull.Value;
            int startIndex = Math.Min(startNumber - 1, bytes.Length);
            return VBScriptByteStrings.GetString(bytes, startIndex, Math.Min(lengthNumber, bytes.Length - startIndex));
        }

        public object INSTRB(object valueToSearch, object valueToSearchFor) => INSTRB(1, valueToSearch, valueToSearchFor);

        public object INSTRB(object startIndex, object valueToSearch, object valueToSearchFor) => INSTRB(startIndex, valueToSearch, valueToSearchFor, VBScriptConstants.vbBinaryCompare);

        /// <summary>
        /// InStrB - the (1-based) byte position of the bytes of valueToSearchFor in valueToSearch from the start (0 if they are not there, Null if
        /// either is Null) - the bytes are compared, also with vbTextCompare (as Windows VBScript does)
        /// </summary>
        public object INSTRB(object startIndex, object valueToSearch, object valueToSearchFor, object compareMode)
        {
            object? startValue = _valueRetriever.VAL(startIndex, "'InStrB'");
            if (startValue == DBNull.Value)
                throw new InvalidUseOfNullException("'InStrB' (start may not be Null)");
            int start = CLNG(startValue, "'InStrB'");
            if (start <= 0)
                throw new InvalidProcedureCallOrArgumentException("'InStrB' (start must be a positive number)");
            _ = CLNG(compareMode, "'InStrB'"); // Windows VBScript compares the bytes whatever the compare mode is

            object? search = _valueRetriever.VAL(valueToSearch, "'InStrB'");
            object? searchFor = _valueRetriever.VAL(valueToSearchFor, "'InStrB'");
            if ((search == DBNull.Value) || (searchFor == DBNull.Value))
                return DBNull.Value;
            byte[] bytes = GetBytesOrNull(search, "'InStrB'")!;
            byte[] bytesToFind = GetBytesOrNull(searchFor, "'InStrB'")!;
            if (bytesToFind.Length == 0)
                return (start <= bytes.Length + 1) ? start : 0;
            if (start > bytes.Length)
                return 0;
            return VBScriptByteStrings.IndexOf(bytes, bytesToFind, start - 1) + 1;
        }

        public object STRREVERSE(object value)
        {
            object? text = _valueRetriever.VAL(value, "'StrReverse'");
            if (text == DBNull.Value)
                throw new InvalidUseOfNullException("'StrReverse'");
            char[] characters = _valueRetriever.STR(text, "'StrReverse'").ToCharArray();
            Array.Reverse(characters);
            return new string(characters);
        }

        public object FILTER(object inputStrings, object value) => FILTER(inputStrings, value, true, VBScriptConstants.vbBinaryCompare);

        public object FILTER(object inputStrings, object value, object include) => FILTER(inputStrings, value, include, VBScriptConstants.vbBinaryCompare);

        /// <summary>
        /// Filter(InputStrings, Value, [Include], [Compare]) - a zero-based array of the strings that contain (or do not contain) the value
        /// </summary>
        public object FILTER(object inputStrings, object value, object include, object compare)
        {
            if (!(inputStrings is Array strings) || (strings.Rank != 1))
                throw new TypeMismatchException("'Filter' (the first argument must be a one-dimensional array)");
            string search = _valueRetriever.STR(value, "'Filter'");
            bool including = CBOOL(include, "'Filter'");
            StringComparison comparison = (CLNG(compare, "'Filter'") == VBScriptConstants.vbTextCompare) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return strings.Cast<object>()
                .Select(item => _valueRetriever.STR(item, "'Filter'"))
                .Where(text => (text.IndexOf(search, comparison) >= 0) == including)
                .Cast<object>()
                .ToArray();
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Formats: FormatNumber / FormatCurrency / FormatPercent(Expression, [NumDigitsAfterDecimal], [IncludeLeadingDigit],
        // [UseParensForNegativeNumbers], [GroupDigits]) - the Tristate arguments are -1 True, 0 False, -2 the regional setting

        public object FORMATNUMBER(object expression) => FormatNumberCore("'FormatNumber'", expression, -1, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Number);
        public object FORMATNUMBER(object expression, object numDigitsAfterDecimal) => FormatNumberCore("'FormatNumber'", expression, numDigitsAfterDecimal, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Number);
        public object FORMATNUMBER(object expression, object numDigitsAfterDecimal, object includeLeadingDigit) => FormatNumberCore("'FormatNumber'", expression, numDigitsAfterDecimal, includeLeadingDigit, VbUseDefault, VbUseDefault, NumberKind.Number);
        public object FORMATNUMBER(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers) => FormatNumberCore("'FormatNumber'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, VbUseDefault, NumberKind.Number);
        public object FORMATNUMBER(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers, object groupDigits) => FormatNumberCore("'FormatNumber'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, groupDigits, NumberKind.Number);

        public object FORMATCURRENCY(object expression) => FormatNumberCore("'FormatCurrency'", expression, -1, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Currency);
        public object FORMATCURRENCY(object expression, object numDigitsAfterDecimal) => FormatNumberCore("'FormatCurrency'", expression, numDigitsAfterDecimal, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Currency);
        public object FORMATCURRENCY(object expression, object numDigitsAfterDecimal, object includeLeadingDigit) => FormatNumberCore("'FormatCurrency'", expression, numDigitsAfterDecimal, includeLeadingDigit, VbUseDefault, VbUseDefault, NumberKind.Currency);
        public object FORMATCURRENCY(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers) => FormatNumberCore("'FormatCurrency'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, VbUseDefault, NumberKind.Currency);
        public object FORMATCURRENCY(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers, object groupDigits) => FormatNumberCore("'FormatCurrency'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, groupDigits, NumberKind.Currency);

        public object FORMATPERCENT(object expression) => FormatNumberCore("'FormatPercent'", expression, -1, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Percent);
        public object FORMATPERCENT(object expression, object numDigitsAfterDecimal) => FormatNumberCore("'FormatPercent'", expression, numDigitsAfterDecimal, VbUseDefault, VbUseDefault, VbUseDefault, NumberKind.Percent);
        public object FORMATPERCENT(object expression, object numDigitsAfterDecimal, object includeLeadingDigit) => FormatNumberCore("'FormatPercent'", expression, numDigitsAfterDecimal, includeLeadingDigit, VbUseDefault, VbUseDefault, NumberKind.Percent);
        public object FORMATPERCENT(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers) => FormatNumberCore("'FormatPercent'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, VbUseDefault, NumberKind.Percent);
        public object FORMATPERCENT(object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers, object groupDigits) => FormatNumberCore("'FormatPercent'", expression, numDigitsAfterDecimal, includeLeadingDigit, useParensForNegativeNumbers, groupDigits, NumberKind.Percent);

        private enum NumberKind
        {
            Number,
            Currency,
            Percent,
        }

        private string FormatNumberCore(string function, object expression, object numDigitsAfterDecimal, object includeLeadingDigit, object useParensForNegativeNumbers, object groupDigits, NumberKind kind)
        {
            object? value = _valueRetriever.VAL(expression, function);
            if (value == DBNull.Value)
                throw new TypeMismatchException(function + " (Null)");
            double number = CDBL_Precise(value, function);
            if (kind == NumberKind.Percent)
                number *= 100;

            NumberFormatInfo format = _culture.NumberFormat;
            int digits = CLNG(numDigitsAfterDecimal, function);
            // the regional setting: the currency digits of the culture; for numbers 2, the Windows setting of (nearly) every locale - .NET's
            // NumberDecimalDigits comes from ICU, which has 3 (eg. for en-US)
            if (digits < 0)
                digits = (kind == NumberKind.Currency) ? format.CurrencyDecimalDigits : 2;
            bool leadingDigit = Tristate(includeLeadingDigit, function, true);
            bool group = Tristate(groupDigits, function, true);
            double rounded = Math.Round(Math.Abs(number), Math.Min(digits, 15), MidpointRounding.AwayFromZero);
            bool negative = (number < 0) && (rounded != 0);

            if (kind == NumberKind.Currency)
            {
                // The currency patterns of the culture (symbol before or after the number); negative amounts in parentheses as the regional setting
                // of Windows has it - .NET's en-US pattern comes from ICU ("-$n"), Windows has "($n)"
                var currency = (NumberFormatInfo)format.Clone();
                if (!group)
                    currency.CurrencyGroupSeparator = "";
                int windowsNegativePattern = (_culture.Name == "en-US") ? 0 : currency.CurrencyNegativePattern;
                bool regionalParens = (windowsNegativePattern == 0) || (windowsNegativePattern == 4) || (windowsNegativePattern == 14) || (windowsNegativePattern == 15);
                bool currencyParens = Tristate(useParensForNegativeNumbers, function, regionalParens);
                currency.CurrencyNegativePattern = (currencyParens == regionalParens) ? windowsNegativePattern : (currencyParens ? 0 : 1);
                string amount = (negative ? -rounded : rounded).ToString("C" + digits.ToString(CultureInfo.InvariantCulture), currency);
                return leadingDigit ? amount : RemoveLeadingZero(amount, currency.CurrencyDecimalSeparator);
            }

            bool parens = Tristate(useParensForNegativeNumbers, function, false);
            string text = rounded.ToString((group ? "N" : "F") + digits.ToString(CultureInfo.InvariantCulture), format);
            if (!leadingDigit)
                text = RemoveLeadingZero(text, format.NumberDecimalSeparator);
            if (kind == NumberKind.Percent)
                text += "%";
            if (negative)
                text = parens ? "(" + text + ")" : format.NegativeSign + text;
            return text;
        }

        /// <summary>
        /// "0.50" -> ".50" (IncludeLeadingDigit False), also inside a currency amount ("$0.50" -> "$.50")
        /// </summary>
        private static string RemoveLeadingZero(string text, string decimalSeparator)
        {
            int index = text.IndexOf("0" + decimalSeparator, StringComparison.Ordinal);
            bool isLeading = (index >= 0) && ((index == 0) || !char.IsDigit(text[index - 1]));
            return isLeading ? text.Remove(index, 1) : text;
        }

        private bool Tristate(object value, string function, bool regionalSetting)
        {
            int tristate = CLNG(value, function);
            return (tristate == VbUseDefault) ? regionalSetting : (tristate != 0);
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Dates

        public object DATEPART(object interval, object valueDate) => DATEPART(interval, valueDate, VBScriptConstants.vbSunday, 1);

        public object DATEPART(object interval, object valueDate, object firstDayOfWeek) => DATEPART(interval, valueDate, firstDayOfWeek, 1);

        /// <summary>
        /// DatePart(Interval, Date, [FirstDayOfWeek], [FirstWeekOfYear]) - yyyy, q, m, y (day of the year), d, w (weekday), ww (week of the year),
        /// h, n, s; an Integer, Null for Null
        /// </summary>
        public object DATEPART(object interval, object valueDate, object firstDayOfWeek, object firstWeekOfYear)
        {
            string intervalText = _valueRetriever.STR(interval, "'DatePart'").ToUpperInvariant();
            object? value = _valueRetriever.VAL(valueDate, "'DatePart'");
            if (value == DBNull.Value)
                return DBNull.Value;
            DateTime date = ToClosestSecond(CDATECore(value, "'DatePart'")); // as Second() (date arithmetic in doubles leaves eg. 29.999 seconds)
            int firstDay = CLNG(firstDayOfWeek, "'DatePart'");
            if (firstDay == 0)
                firstDay = (int)_culture.DateTimeFormat.FirstDayOfWeek + 1; // vbUseSystemDayOfWeek
            if ((firstDay < 1) || (firstDay > 7))
                throw new InvalidProcedureCallOrArgumentException($"'DatePart' (FirstDayOfWeek {firstDay})");
            int firstWeek = CLNG(firstWeekOfYear, "'DatePart'");
            switch (intervalText)
            {
                case "YYYY": return (short)date.Year;
                case "Q": return (short)(((date.Month - 1) / 3) + 1);
                case "M": return (short)date.Month;
                case "Y": return (short)date.DayOfYear;
                case "D": return (short)date.Day;
                case "W": return (short)(((((int)date.DayOfWeek - (firstDay - 1)) + 7) % 7) + 1);
                case "WW":
                    CalendarWeekRule rule;
                    switch (firstWeek)
                    {
                        case 0:
                        case 1: rule = CalendarWeekRule.FirstDay; break; // vbUseSystem / vbFirstJan1
                        case 2: rule = CalendarWeekRule.FirstFourDayWeek; break; // vbFirstFourDays
                        case 3: rule = CalendarWeekRule.FirstFullWeek; break; // vbFirstFullWeek
                        default: throw new InvalidProcedureCallOrArgumentException($"'DatePart' (FirstWeekOfYear {firstWeek})");
                    }
                    return (short)CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(date, rule, (DayOfWeek)(firstDay - 1));
                case "H": return (short)date.Hour;
                case "N": return (short)date.Minute;
                case "S": return (short)date.Second;
                default:
                    throw new InvalidProcedureCallOrArgumentException($"'DatePart' (unknown interval '{intervalText}')");
            }
        }

        public object MONTHNAME(object value) => MONTHNAME(value, false);

        /// <summary>
        /// MonthName(Month, [Abbreviate]) - the name of the month in the language of the culture
        /// </summary>
        public object MONTHNAME(object value, object abbreviate)
        {
            object? month = _valueRetriever.VAL(value, "'MonthName'");
            if (month == DBNull.Value)
                throw new InvalidUseOfNullException("'MonthName'");
            int number = CLNG(month, "'MonthName'");
            if ((number < 1) || (number > 12))
                throw new InvalidProcedureCallOrArgumentException($"'MonthName' ({number} is not a month)");
            return CBOOL(abbreviate, "'MonthName'") ? _culture.DateTimeFormat.GetAbbreviatedMonthName(number) : _culture.DateTimeFormat.GetMonthName(number);
        }

        // ---------------------------------------------------------------------------------------------------------------------------------------
        // Script engine

        public string SCRIPTENGINE() => "VBScript";

        /// <summary>
        /// The version of the VBScript engine of current Windows versions (ScriptEngineMajorVersion 10, ScriptEngineMinorVersion 8)
        /// </summary>
        public int SCRIPTENGINEMAJORVERSION() => 10;

        public int SCRIPTENGINEMINORVERSION() => 8;

        public int SCRIPTENGINEBUILDVERSION() => 16384;

        /// <summary>
        /// GetLocale - the locale identifier (LCID) of the culture that the script runs with
        /// </summary>
        public object GETLOCALE() => _culture.LCID;
    }
}
