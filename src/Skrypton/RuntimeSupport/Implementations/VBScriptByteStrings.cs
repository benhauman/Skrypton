using System;
using System.Collections.Generic;

namespace Skrypton.RuntimeSupport.Implementations
{
    /// <summary>
    /// VBScript's byte strings (LenB, LeftB, RightB, MidB, InStrB, AscB, ChrB - issue #93): a VBScript string is its UTF-16LE bytes and may have
    /// an odd number of them ("ChrB(65)" is one byte, "ChrB(65) &amp; ChrB(0)" is "A"). A .NET string has whole characters, so a single byte is
    /// kept as a lone low surrogate (U+DC00 + the byte), which does not occur in text (a low surrogate in text follows a high surrogate).
    /// </summary>
    internal static class VBScriptByteStrings
    {
        private const char SingleByteBase = '\uDC00';

        /// <summary>
        /// Is the character at the index a single byte (a lone low surrogate in U+DC00..U+DCFF)
        /// </summary>
        public static bool IsSingleByte(string value, int index)
        {
            char c = value[index];
            if ((c < '\uDC00') || (c > '\uDCFF'))
                return false;
            return (index == 0) || !char.IsHighSurrogate(value[index - 1]);
        }

        /// <summary>
        /// Does the string end with a single byte (so that the bytes of a string appended to it have to be joined with it)
        /// </summary>
        public static bool EndsWithSingleByte(string value) => (value.Length > 0) && IsSingleByte(value, value.Length - 1);

        public static byte[] GetBytes(string value)
        {
            var bytes = new List<byte>(value.Length * 2);
            for (int i = 0; i < value.Length; i++)
            {
                if (IsSingleByte(value, i))
                {
                    bytes.Add((byte)(value[i] & 0xFF));
                    continue;
                }
                bytes.Add((byte)(value[i] & 0xFF));
                bytes.Add((byte)(value[i] >> 8));
            }
            return bytes.ToArray();
        }

        public static string GetString(byte[] bytes) => GetString(bytes, 0, bytes.Length);

        /// <summary>
        /// The string of the bytes - two bytes per character (little endian), an odd last byte as a single byte
        /// </summary>
        public static string GetString(byte[] bytes, int start, int count)
        {
            var text = new char[(count + 1) / 2];
            for (int i = 0; i < count / 2; i++)
                text[i] = (char)(bytes[start + (2 * i)] | (bytes[start + (2 * i) + 1] << 8));
            if ((count % 2) == 1)
                text[text.Length - 1] = (char)(SingleByteBase + bytes[start + count - 1]);
            return new string(text);
        }

        /// <summary>
        /// l &amp; r when l ends with a single byte: the bytes of r follow that byte
        /// </summary>
        public static string Concat(string l, string r)
        {
            byte[] left = GetBytes(l);
            byte[] right = GetBytes(r);
            var bytes = new byte[left.Length + right.Length];
            Array.Copy(left, bytes, left.Length);
            Array.Copy(right, 0, bytes, left.Length, right.Length);
            return GetString(bytes);
        }

        /// <summary>
        /// The 0-based index of the bytes in the bytes from the start index (-1 if they are not there)
        /// </summary>
        public static int IndexOf(byte[] bytes, byte[] search, int start)
        {
            for (int i = start; i + search.Length <= bytes.Length; i++)
            {
                int matched = 0;
                while ((matched < search.Length) && (bytes[i + matched] == search[matched]))
                    matched++;
                if (matched == search.Length)
                    return i;
            }
            return -1;
        }
    }
}
