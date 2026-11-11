using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace Skrypton.RuntimeSupport
{
    /// <summary>
    /// Creates NotSupportedExceptions whose message says where the problem is: the VBScript line and source (when known) and the Skrypton source
    /// location that raised it (file, line and member, taken from the caller attributes), so that an unsupported construct can be located quickly
    /// </summary>
    internal static class NotSupportedExceptionFactory
    {
        /// <param name="message">What is not supported</param>
        /// <param name="vbsLineIndex">The zero-based VBScript line index, if known</param>
        /// <param name="vbsSource">The VBScript source (or rendered expression) involved, if known</param>
        public static NotSupportedException Create(
            string message,
            int? vbsLineIndex = null,
            string? vbsSource = null,
            [CallerMemberName] string callerMemberName = "",
            [CallerFilePath] string callerFilePath = "",
            [CallerLineNumber] int callerLineNumber = 0)
        {
            return new NotSupportedException(GetMessage(message, vbsLineIndex, vbsSource, callerMemberName, callerFilePath, callerLineNumber));
        }

        internal static string GetMessage(string message, int? vbsLineIndex, string? vbsSource, string callerMemberName, string callerFilePath, int callerLineNumber)
        {
            var text = new StringBuilder(string.IsNullOrWhiteSpace(message) ? "Not supported" : message.Trim());
            if (vbsLineIndex.HasValue && (vbsLineIndex.Value >= 0))
                text.Append(" at line ").Append(vbsLineIndex.Value + 1);
            if (!string.IsNullOrWhiteSpace(vbsSource))
                text.Append(vbsLineIndex.HasValue ? ": '" : " in '").Append(vbsSource!.Trim()).Append('\'');

            // The file name only (CallerFilePath is the full path on the build machine)
            string fileName = string.IsNullOrEmpty(callerFilePath) ? "" : Path.GetFileName(callerFilePath.Replace('\\', '/'));
            text.Append(" [").Append(fileName).Append(':').Append(callerLineNumber);
            if (!string.IsNullOrEmpty(callerMemberName))
                text.Append(' ').Append(callerMemberName);
            text.Append(']');
            return text.ToString();
        }
    }
}
