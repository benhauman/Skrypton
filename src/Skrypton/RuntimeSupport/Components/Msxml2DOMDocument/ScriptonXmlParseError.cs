using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// DOMDocument.parseError (IXMLDOMParseError - issue #173): why the most recent load / loadXML failed - the errorCode and reason of MSXML for
    /// the errors that scripts meet most (unclosed tags, no / several top level elements, mismatched end tag, duplicate attribute, undefined entity,
    /// unquoted attribute value, a missing file), the line, the position in the line (1-based), the offset in the text (filepos) and the text of the
    /// line (srcText). An error at the end of the text (unclosed tags, no top level element) has line and position 0, as with MSXML.
    /// </summary>
    [SourceClassName("IXMLDOMParseError")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonXmlParseError : ScriptonIReflectOnClrType
    {
        public static readonly ScriptonXmlParseError None = new ScriptonXmlParseError(0, "", "", 0, 0, 0, "");

        private ScriptonXmlParseError(int errorCode, string reason, string srcText, int line, int linepos, int filepos, string url)
            : base(typeof(ScriptonXmlParseError))
        {
            this.errorCode = errorCode;
            this.reason = reason;
            this.srcText = srcText;
            this.line = line;
            this.linepos = linepos;
            this.filepos = filepos;
            this.url = url;
        }

        [DispId(0)]
        public int errorCode { get; }

        [DispId(0xB3)]
        public string url { get; }

        [DispId(0xB4)]
        public string reason { get; }

        [DispId(0xB5)]
        public string srcText { get; }

        [DispId(0xB6)]
        public int line { get; }

        [DispId(0xB7)]
        public int linepos { get; }

        [DispId(0xB8)]
        public int filepos { get; }

        /// <summary>
        /// The MSXML error for the .NET error of the load: the codes and texts of MSXML (verified with cscript), otherwise E_FAIL with the text of .NET
        /// </summary>
        internal static ScriptonXmlParseError From(Exception error, string source, string url)
        {
            switch (error)
            {
                case System.IO.DirectoryNotFoundException _:
                    return new ScriptonXmlParseError(unchecked((int)0x80070003), WindowsMessage(3), "", 0, 0, 0, url);
                case System.IO.FileNotFoundException _:
                    return new ScriptonXmlParseError(unchecked((int)0x80070002), WindowsMessage(2), "", 0, 0, 0, url);
                case XmlException xml:
                    string message = Regex.Replace(xml.Message, @"\s*Line \d+, position \d+\.\s*$", "");
                    Match match;
                    if ((match = Regex.Match(message, @"^Unexpected end of file has occurred\. The following elements are not closed: (.+?)\.?$")).Success)
                        return AtEnd(0xC00CE553, "The following tags were not closed: " + string.Join(", ", match.Groups[1].Value.TrimEnd('.').Split(',').Select(t => t.Trim()).Reverse()) + ".", url); // MSXML names the outermost first
                    if (message.StartsWith("Root element is missing", StringComparison.Ordinal))
                        return AtEnd(0xC00CE558, "XML document must have a top level element.", url);
                    if ((match = Regex.Match(message, @"^The '(.+?)' start tag on line \d+ position \d+ does not match the end tag of '(.+?)'")).Success)
                        return At(xml, 0xC00CE56D, $"End tag '{match.Groups[2].Value}' does not match the start tag '{match.Groups[1].Value}'.", source, url);
                    if (Regex.IsMatch(message, @"^'.+?' is a duplicate attribute name"))
                        return At(xml, 0xC00CE554, "Duplicate attribute.", source, url, afterTheTag: true); // MSXML reports it after the start tag
                    if (message.StartsWith("Data at the root level is invalid", StringComparison.Ordinal))
                        return At(xml, 0xC00CE556, "Invalid at the top level of the document.", source, url);
                    if (message.StartsWith("There are multiple root elements", StringComparison.Ordinal))
                        return At(xml, 0xC00CE555, "Only one top level element is allowed in an XML document.", source, url);
                    if ((match = Regex.Match(message, @"^Reference to undeclared entity '(.+?)'")).Success)
                        return At(xml, 0xC00CE002, $"Reference to undefined entity '{match.Groups[1].Value}'.", source, url);
                    if (Regex.IsMatch(message, @"is an unexpected token\. The expected token is '""' or '''"))
                        return At(xml, 0xC00CE502, "A string literal was expected, but no opening quote character was found.", source, url);
                    return At(xml, 0x80004005, message, source, url);
                default:
                    return new ScriptonXmlParseError(unchecked((int)0x80004005), error.Message + "\r\n", "", 0, 0, 0, url);
            }
        }

        private static ScriptonXmlParseError AtEnd(uint errorCode, string reason, string url) => new ScriptonXmlParseError(unchecked((int)errorCode), reason + "\r\n", "", 0, 0, 0, url);

        private static ScriptonXmlParseError At(XmlException error, uint errorCode, string reason, string source, string url, bool afterTheTag = false)
        {
            int line = Math.Max(0, error.LineNumber), linepos = Math.Max(0, error.LinePosition);
            string[] lines = Regex.Split(source ?? "", "\r\n|\n|\r");
            string srcText = ((line >= 1) && (line <= lines.Length)) ? lines[line - 1] : "";
            if (afterTheTag && (linepos >= 1))
            {
                int tagEnd = srcText.IndexOf('>', Math.Min(linepos - 1, srcText.Length));
                if (tagEnd >= 0)
                    linepos = tagEnd + 2; // the position after the ">" of the tag
            }
            int filepos = 0;
            if (line >= 1)
            {
                // the offset of the position in the text (the line breaks counted as they are)
                int lineStart = 0;
                for (int current = 1; (current < line) && (lineStart < (source ?? "").Length); current++)
                {
                    Match lineBreak = Regex.Match(source!.Substring(lineStart), "\r\n|\n|\r");
                    lineStart = lineBreak.Success ? (lineStart + lineBreak.Index + lineBreak.Length) : source.Length;
                }
                filepos = lineStart + Math.Max(0, linepos - 1);
            }
            return new ScriptonXmlParseError(unchecked((int)errorCode), reason + "\r\n", srcText, line, linepos, filepos, url);
        }

        private static string WindowsMessage(int win32Error)
        {
            string message = new System.ComponentModel.Win32Exception(win32Error).Message.TrimEnd();
            return (message.EndsWith(".", StringComparison.Ordinal) ? message : message + ".") + "\r\n"; // .NET Framework removes the period
        }
    }
}
