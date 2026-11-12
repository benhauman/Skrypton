using System;
using System.Globalization;
using Skrypton.RuntimeSupport.Exceptions;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The conversion of an argument to the type of the parameter of a built-in member, as COM converts the argument of a call to the type of the
    /// parameter (VariantChangeType, issue #138): a number or "True" / "False" for a Boolean (not 0 is True), a Boolean (-1 / 0), a number or a
    /// numeric string for a number (rounded to even for an integer), a number or a Boolean for a string. A value that can not be converted is a
    /// Type mismatch (13), a number out of the range of the type an Overflow (6).
    /// </summary>
    internal static class VBScriptArgumentConversion
    {
        public static object? ChangeType(object? value, Type type, CultureInfo culture)
        {
            if (value is DBNull)
                throw new TypeMismatchException(null, null); // Null converts to no type
            Type target = Nullable.GetUnderlyingType(type) ?? type;
            if (value == null) // Empty: False, 0 or ""
                return (target == typeof(string)) ? "" : (target.IsValueType ? Activator.CreateInstance(target) : null);
            if (target.IsInstanceOfType(value))
                return value;
            if (target.IsEnum)
                return Enum.ToObject(target, ChangeType(value, Enum.GetUnderlyingType(target), culture)!);
            if (target == typeof(bool))
                return ToBoolean(value, culture);
            if (target == typeof(string))
                return ToText(value, culture);
            if (IsInteger(target))
                return ToInteger(ToNumber(value, culture), target);
            if ((target == typeof(double)) || (target == typeof(float)) || (target == typeof(decimal)))
                return ToFloatingPoint(ToNumber(value, culture), target);
            throw new TypeMismatchException(null, null);
        }

        private static bool ToBoolean(object value, CultureInfo culture)
        {
            if (value is string text)
            {
                if (string.Equals(text.Trim(), "True", StringComparison.OrdinalIgnoreCase))
                    return true;
                if (string.Equals(text.Trim(), "False", StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return ToNumber(value, culture) != 0;
        }

        private static string ToText(object value, CultureInfo culture)
        {
            switch (value)
            {
                case bool b:
                    return b ? "True" : "False";
                case double d:
                    return d.ToString("G15", culture); // VBScript shows 15 significant digits (a Single 7)
                case float f:
                    return f.ToString("G7", culture);
                case byte _:
                case short _:
                case int _:
                case long _:
                case decimal _:
                    return Convert.ToString(value, culture)!;
                default:
                    throw new TypeMismatchException(null, null);
            }
        }

        private static double ToNumber(object value, CultureInfo culture)
        {
            switch (value)
            {
                case bool b:
                    return b ? -1 : 0;
                case string text:
                    if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, culture, out double parsed))
                        return parsed;
                    throw new TypeMismatchException(null, null);
                case byte _:
                case sbyte _:
                case short _:
                case ushort _:
                case int _:
                case uint _:
                case long _:
                case ulong _:
                case float _:
                case double _:
                case decimal _:
                    return Convert.ToDouble(value, CultureInfo.InvariantCulture);
                default:
                    throw new TypeMismatchException(null, null);
            }
        }

        private static bool IsInteger(Type type) =>
            (type == typeof(byte)) || (type == typeof(sbyte)) || (type == typeof(short)) || (type == typeof(ushort)) ||
            (type == typeof(int)) || (type == typeof(uint)) || (type == typeof(long)) || (type == typeof(ulong));

        private static object ToInteger(double number, Type type)
        {
            double rounded = Math.Round(number, MidpointRounding.ToEven);
            try
            {
                return Convert.ChangeType(checked((decimal)rounded), type, CultureInfo.InvariantCulture);
            }
            catch (OverflowException e)
            {
                throw new VBScriptOverflowException(number, e);
            }
        }

        private static object ToFloatingPoint(double number, Type type)
        {
            try
            {
                return Convert.ChangeType(number, type, CultureInfo.InvariantCulture);
            }
            catch (OverflowException e)
            {
                throw new VBScriptOverflowException(number, e);
            }
        }
    }
}
