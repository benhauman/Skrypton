using System;
using System.Diagnostics;
using System.Text;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("{_content}")]
    public class TranslatedStatement
    {
        public TranslatedStatement(int lineIndexOfStatementStartInSource) // 'empty' ctor
            : this(true, "", 0, lineIndexOfStatementStartInSource)
        {
        }
        public TranslatedStatement(string content, int indentationDepth, int lineIndexOfStatementStartInSource)
            : this(false, content, indentationDepth, lineIndexOfStatementStartInSource)
        {
        }
        private TranslatedStatement(bool isEmpty, string content, int indentationDepth, int lineIndexOfStatementStartInSource)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (content != content.Trim())
                throw new ArgumentException("content may be blank but may not have any leading or trailing whitespace");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");
            if (lineIndexOfStatementStartInSource < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndexOfStatementStartInSource), "must be zero or greater");

            if (!isEmpty && content.Length == 0)
                throw new InvalidOperationException("Use the 'empty' ctor.");
            _content = content;
            _indentationDepth = indentationDepth;
            LineIndexOfStatementStartInSource = lineIndexOfStatementStartInSource;

        }

        /// <summary>
        /// This will never be null, though it may be blank if it represents a blank line. It will never have any leading or trailing whitespace.
        /// </summary>
        private readonly string _content;
        internal bool HasContent => _content.Length > 0;

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        private readonly int _indentationDepth;

        /// <summary>
        /// This will indicate where in the VBScript source that code exists that resulted in the current line of C# being generated. Not all lines of C# have
        /// a direct source in VBScript and some lines may relate to multiple lines of VBScript (particularly if the VBScript lines were split up using the
        /// line continuation character). As such, there are times when this value will be somewhat approximate (and blank lines often have a value of
        /// zero, since they are not of any significant importance). This value will always be zero or greater.
        /// </summary>
        public int LineIndexOfStatementStartInSource { get; }

        private string? _inlineCommentIfAny;
        internal void AppendInlineComment(string translatedCommentContent)
        {
            if (_inlineCommentIfAny != null)
                throw new InvalidOperationException("Inline comment already set;");
            _inlineCommentIfAny = translatedCommentContent;
        }

        public StringBuilder RenderTranslatedStatement(StringBuilder tb)
        {
            if (tb == null) throw new ArgumentNullException(nameof(tb));
            if (HasContent || _inlineCommentIfAny != null)
            {
                if (_indentationDepth > 0)
                {
                    _ = tb.Append(new string(' ', _indentationDepth * 4));
                }

                tb.Append(_content);
                if (_inlineCommentIfAny != null)
                {
                    tb.Append(" //").Append(_inlineCommentIfAny);
                }
            }
            else
            {
                //tb.Append(s._content); // no indention for blank lines
            }
            return tb;
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

    public enum TranslatedStatementKind
    {
        Unknown
    }
}
