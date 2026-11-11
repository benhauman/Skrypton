using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// The column of the statement that raised an error (issue #87): VBScript reports the column where the statement starts. A line with one
    /// statement has its column; on a line with more statements ("Dim o : Set o = Nothing : o.Go") the statement is the one whose names appear
    /// in the C# line that failed.
    /// </summary>
    internal static class ScriptStatementColumns
    {
        private static readonly Regex Identifier = new Regex(@"[A-Za-z_][A-Za-z0-9_]*", RegexOptions.CultureInvariant);

        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Dim", "Set", "Let", "Call", "If", "Then", "Else", "ElseIf", "End", "And", "Or", "Not", "Xor", "Eqv", "Imp", "Mod", "Is", "Nothing",
            "Null", "Empty", "True", "False", "New", "On", "Error", "Resume", "Next", "GoTo", "For", "To", "Step", "Each", "In", "Do", "Loop",
            "While", "Until", "Wend", "Sub", "Function", "Exit", "ReDim", "Preserve", "Const", "Option", "Explicit", "With", "Select", "Case",
            "Public", "Private", "ByVal", "ByRef", "Property", "Get", "Class", "Erase",
        };

        /// <summary>
        /// The 0-based column of the statement on the VBScript line (0 if it cannot be determined) - variableName: the statement must contain it
        /// </summary>
        public static int Find(string scriptLine, string csharpLine, string? variableName)
        {
            List<(int Column, string Text)> statements = Split(scriptLine ?? "");
            if (statements.Count == 0)
                return 0;
            if (statements.Count == 1)
                return statements[0].Column;

            // declarations (Dim, Const, Sub, End Sub, ..) do not run, so they do not raise runtime errors
            List<(int Column, string Text)> executable = statements.Where(s => !IsDeclaration(s.Text)).ToList();
            if (executable.Count > 0)
                statements = executable;
            if (statements.Count == 1)
                return statements[0].Column;

            if (!string.IsNullOrEmpty(variableName))
            {
                foreach ((int column, string text) in statements)
                {
                    if (Identifier.Matches(text).Cast<Match>().Any(m => string.Equals(m.Value, variableName, StringComparison.OrdinalIgnoreCase)))
                        return column;
                }
            }

            // the statement whose names (variables, members, string literals) are in the failing C# line
            int bestColumn = statements[0].Column;
            int bestScore = 0;
            foreach ((int column, string text) in statements)
            {
                int score = Identifier.Matches(text).Cast<Match>()
                    .Select(m => m.Value)
                    .Where(name => !Keywords.Contains(name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count(name => Regex.IsMatch(csharpLine ?? "", @"\b" + Regex.Escape(name) + @"\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
                if (score > bestScore)
                {
                    bestScore = score;
                    bestColumn = column;
                }
            }
            return bestColumn;
        }

        /// <summary>
        /// The statements of the line - separated by ':' outside of strings, up to a comment - with their 0-based columns
        /// </summary>
        internal static List<(int Column, string Text)> Split(string line)
        {
            var statements = new List<(int, string)>();
            int start = 0;
            bool inString = false;
            for (int i = 0; i <= line.Length; i++)
            {
                char c = (i < line.Length) ? line[i] : ':';
                if (inString)
                {
                    if (c == '"')
                        inString = false; // "" is an escaped quote: it ends and starts the string again
                    continue;
                }
                if (c == '"')
                {
                    inString = true;
                    continue;
                }
                bool comment = (c == '\'') || IsRemAt(line, i, start);
                if ((c == ':') || comment)
                {
                    string text = line.Substring(start, i - start);
                    int offset = text.Length - text.TrimStart().Length;
                    if (text.Trim().Length > 0)
                        statements.Add((start + offset, text.Trim()));
                    if (comment)
                        break;
                    start = i + 1;
                }
            }
            return statements;
        }

        private static readonly Regex Declaration = new Regex(
            @"^(Dim|Const|Option|Class|End\s+(Sub|Function|Property|Class)|((Public|Private)\s+)?(Default\s+)?(Sub|Function|Property))\b|^(Public|Private)\s+[A-Za-z_]",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static bool IsDeclaration(string statement) => Declaration.IsMatch(statement);

        private static bool IsRemAt(string line, int index, int statementStart)
        {
            // "Rem" as the first word of a statement is a comment
            if ((index + 3 > line.Length) || !string.Equals(line.Substring(index, 3), "Rem", StringComparison.OrdinalIgnoreCase))
                return false;
            bool atStatementStart = line.Substring(statementStart, index - statementStart).Trim().Length == 0;
            bool wordEnds = (index + 3 == line.Length) || char.IsWhiteSpace(line[index + 3]);
            return atStatementStart && wordEnds;
        }
    }
}
