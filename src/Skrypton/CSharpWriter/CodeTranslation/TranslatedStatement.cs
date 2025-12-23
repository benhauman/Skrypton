using System;
using System.Diagnostics;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("{Content}")]
    public class TranslatedStatement
    {
        public TranslatedStatement(string content, int indentationDepth, int lineIndexOfStatementStartInSource)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (content != content.Trim())
                throw new ArgumentException("content may be blank but may not have any leading or trailing whitespace");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");
            if (lineIndexOfStatementStartInSource < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndexOfStatementStartInSource), "must be zero or greater");

            Content = content;
            IndentationDepth = indentationDepth;
            LineIndexOfStatementStartInSource = lineIndexOfStatementStartInSource;

        }

        /// <summary>
        /// This will never be null, though it may be blank if it represents a blank line. It will never have any leading or trailing whitespace.
        /// </summary>
        public string Content { get; private set; }
        public bool HasContent => Content != "";

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        public int IndentationDepth { get; private set; }

        /// <summary>
        /// This will indicate where in the VBScript source that code exists that resulted in the current line of C# being generated. Not all lines of C# have
        /// a direct source in VBScript and some lines may relate to multiple lines of VBScript (particularly if the VBScript lines were split up using the
        /// line continuation character). As such, there are times when this value will be somewhat approximate (and blank lines often have a value of
        /// zero, since they are not of any significant importance). This value will always be zero or greater.
        /// </summary>
        public int LineIndexOfStatementStartInSource { get; private set; }

        internal void AppendInlineComment(string translatedCommentContent)
        {
            if (Content != "")
            {
                Content += " ";
            }
            Content += translatedCommentContent;
        }
    }

    public sealed class TranslatedVariableDeclarationStatement : TranslatedStatement
    {
        public TranslatedVariableDeclarationStatement(string variableAccessToken, string content, int indentationDepth, int lineIndexOfStatementStartInSource)
            : base(content, indentationDepth, lineIndexOfStatementStartInSource)
        {
            if (string.IsNullOrEmpty(variableAccessToken))
                throw new ArgumentException("Value cannot be null or empty.", nameof(variableAccessToken));
            VariableAccessToken = variableAccessToken;
        }
        public string VariableAccessToken { get; private set; }
    }
}
