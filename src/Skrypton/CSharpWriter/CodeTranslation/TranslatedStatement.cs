using Skrypton.CSharpWriter.Lists;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("{_content}")]
    public class TranslatedStatement
    {
        public TranslatedStatement(TranslatedStatementKind kind, int lineIndexOfStatementStartInSource) // 'empty' ctor
            : this(true, kind, true, "", 0, lineIndexOfStatementStartInSource)
        {
        }
        public TranslatedStatement(TranslatedStatementKind kind, string content, int indentationDepth, int lineIndexOfStatementStartInSource)
            : this(false, kind, false, content, indentationDepth, lineIndexOfStatementStartInSource)
        {
        }
        internal TranslatedStatement(bool isStatement, TranslatedStatementKind kind, string content, int indentationDepth, int lineIndexOfStatementStartInSource)
            : this(true, kind, false, content, indentationDepth, lineIndexOfStatementStartInSource)
        {
        }
        private TranslatedStatement(bool isStatement, TranslatedStatementKind kind, bool isEmpty, string content, int indentationDepth, int lineIndexOfStatementStartInSource)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (!isStatement && content != content.Trim())
                throw new ArgumentException("content may be blank but may not have any leading or trailing whitespace");
            if (indentationDepth < 0)
                throw new ArgumentOutOfRangeException(nameof(indentationDepth), "must be zero or greater");
            if (lineIndexOfStatementStartInSource < 0)
                throw new ArgumentOutOfRangeException(nameof(lineIndexOfStatementStartInSource), "must be zero or greater");

            if (!isStatement && !isEmpty && content.Length == 0)
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

        internal const char NewLineNormalized = '\n'; // 10: line feed (LF) character.
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
            : base(TranslatedStatementKind.VariableDeclarationStatement, content, indentationDepth, lineIndexOfStatementStartInSource)
        {
            if (string.IsNullOrEmpty(variableAccessToken))
                throw new ArgumentException("Value cannot be null or empty.", nameof(variableAccessToken));
            VariableAccessToken = variableAccessToken;
        }
        public string VariableAccessToken { get; private set; }
    }

    public enum TranslatedStatementKind
    {
        Unknown,
        RawText,
        Comment,
        CurlyBraceOpen,
        CurlyBraceClose,
        Else,
        IfWithCondition, // if ()
        IfText,
        TryBegin,
        FinallyClause,
        NamespaceBegin,
        SetText,
        SupportHandleError,
        VariableDeclarationStatement,
        PropertyDeclarationStatement,
        FieldDeclarationStatement,
        ReturnText,
        UsingText,
        ////////////////////
        OutermostCodeText,
        ClassDeclarationStatement,
        ConstructorDeclarationStatement,
    }

    internal static class TranslatedStatementFactory
    {
        internal static CSharpStatementBuilderConstructor CreateConstructor(int indentationDepth, int lineIndexOfStatementStartInSource) => CSharpCodeBuilder.Init(new CSharpStatementBuilderConstructor(), indentationDepth, lineIndexOfStatementStartInSource);
        internal static CSharpClassBuilder CreateClass(int indentationDepth, int lineIndexOfStatementStartInSource) => CSharpCodeBuilder.Init(new CSharpClassBuilder(), indentationDepth, lineIndexOfStatementStartInSource);

        internal static CSharpCodeBuilder FromRawText(string rawText, int indentationDepth, int lineIndexOfStatementStartInSource) => CSharpCodeBuilder.Init(new CSharpCodeBuilderRawText(rawText), indentationDepth, lineIndexOfStatementStartInSource);
    }

    internal abstract class CSharpCodeBuilder
    {
        private readonly TranslatedStatementKind _kind;
        protected int IndentationDepth { get; private set; }
        protected int LineIndexOfStatementStartInSource { get; private set; }

        protected string IndentationSpace { get; private set; } = "";

        private readonly List<CSharpCodeBuilder> _builders = new List<CSharpCodeBuilder>();

        protected CSharpCodeBuilder(TranslatedStatementKind kind)
        {
            _kind = kind;
        }
        internal static TBuilder Init<TBuilder>(TBuilder builder, int indentationDepth, int lineIndexOfStatementStartInSource) where TBuilder : CSharpCodeBuilder
        {
            builder.IndentationDepth = indentationDepth;
            builder.LineIndexOfStatementStartInSource = lineIndexOfStatementStartInSource;
            builder.IndentationSpace = indentationDepth == 0 ? "" : new string(' ', indentationDepth * 4);
            return builder;
        }
        protected void AddChildBuilder(CSharpCodeBuilder builder)
        {
            _builders.Add(builder);
        }
        protected IReadOnlyCollection<CSharpCodeBuilder> ChildBuilders() => _builders;

        internal TranslatedStatement BuildTranslatedStatement()
        {
            StringBuilder sb = new StringBuilder();
            DoBuildTranslatedStatement(sb);
            return new TranslatedStatement(true, _kind, sb.ToString(), 0, LineIndexOfStatementStartInSource);
        }
        internal void RenderTranslatedStatement(StringBuilder sb)
        {
            DoBuildTranslatedStatement(sb);
        }
        protected abstract void DoBuildTranslatedStatement(StringBuilder tb);

        internal virtual bool HasContent => true;

        protected const char NewLineNormalized = '\n';

    }

    internal sealed class CSharpParameterDeclaration
    {
        public string ParameterTypeName { get; }
        public string ParameterName { get; }

        public CSharpParameterDeclaration(string parameterTypeName, string parameterName)
        {
            ParameterTypeName = parameterTypeName;
            ParameterName = parameterName;
        }
    }

    internal sealed class CSharpStatementBuilderConstructor() : CSharpCodeBuilder(TranslatedStatementKind.ConstructorDeclarationStatement)
    {
        private string? _className;
        private readonly List<CSharpParameterDeclaration> _parameters = new List<CSharpParameterDeclaration>();
        private readonly List<string> _baseParameters = new List<string>();
        public CSharpStatementBuilderConstructor ClassName(string className)
        {
            _className = className;
            return this;
        }
        public CSharpStatementBuilderConstructor Parameter(string parameterTypeName, string parameterName, bool asBaseParameter)
        {
            _parameters.Add(new CSharpParameterDeclaration(parameterTypeName, parameterName));
            if (asBaseParameter)
            {
                _baseParameters.Add(parameterName);
            }
            return this;
        }

        private readonly List<TranslatedStatement> _bodyStatements = new List<TranslatedStatement>();
        public CSharpStatementBuilderConstructor AddStatement(TranslatedStatement stmt)
        {
            _bodyStatements.Add(stmt);
            return this;
        }

        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            string indentationCtor = new string(' ', IndentationDepth * 4);
            string indentationBody = new string(' ', IndentationDepth * 6);

            tb.Append(indentationCtor).Append("public").Append(' ').Append(_className).Append('(');
            for (int ixPrm = 0; ixPrm < _parameters.Count; ixPrm++)
            {
                var prm = _parameters[ixPrm];
                if (ixPrm > 0)
                    tb.Append(", ");
                tb.Append(prm.ParameterTypeName).Append(' ').Append(prm.ParameterName);
            }
            tb.Append(')');
            if (_baseParameters.Count > 0)
            {
                tb.Append(" : base(");
                //  : base({parameter1Name})"
                for (int ixPrm = 0; ixPrm < _baseParameters.Count; ixPrm++)
                {
                    var prmParameterName = _baseParameters[ixPrm];
                    if (ixPrm > 0)
                        tb.Append(", ");
                    tb.Append(prmParameterName);
                }
                tb.Append(')');
            }
            tb.Append(TranslatedStatement.NewLineNormalized);

            // body
            tb.Append(indentationCtor).Append('{').Append(TranslatedStatement.NewLineNormalized);

            foreach (var bodyStatement in _bodyStatements)
            {
                tb.Append(indentationBody);
                bodyStatement.RenderTranslatedStatement(tb);
                tb.Append(TranslatedStatement.NewLineNormalized);
            }

            tb.Append(indentationCtor).Append('}');//.AppendLine()
        }
    }

    internal sealed class CSharpOutermostCodeBuilder() : CSharpCodeBuilder(TranslatedStatementKind.OutermostCodeText)
    {
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            throw new NotImplementedException();
        }


        internal CSharpOutermostCodeBuilder AddBuilder(CSharpCodeBuilder builder)
        {
            base.AddChildBuilder(builder); return this;
        }
        public CSharpOutermostCodeBuilder AddRange(IReadOnlyCollection<TranslatedStatement> values)
        {
            foreach (TranslatedStatement value in values)
            {
                AddChildBuilder(new CSharpCodeBuilderWrap(value));
            }
            return this;
        }
        public CSharpOutermostCodeBuilder Add(TranslatedStatement value)
        {
            AddChildBuilder(new CSharpCodeBuilderWrap(value));
            return this;
        }

        private List<CSharpCodeBuilder> RemoveRunsOfBlankLines()
        {
            List<CSharpCodeBuilder> children = ChildBuilders().ToList();

            return children
                .Select((s, i) => ((i == 0) || (s.HasContent) || (children[i - 1].HasContent)) ? s : null)
                .Where(s => s != null)
                .Select(s => s!)
                .ToList();
        }

        internal string RenderTranslatedProgramCode()
        {
            var children = RemoveRunsOfBlankLines();

            StringBuilder tb = new StringBuilder();
            foreach (CSharpCodeBuilder s in children)
            {
                s.RenderTranslatedStatement(tb);
                tb.Append(NewLineNormalized);
            }

            string csText = tb.ToString();
            return csText;
        }
    }

    internal sealed class CSharpCodeBuilderWrap(TranslatedStatement statement) : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        private readonly TranslatedStatement _statement = statement;

        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            _statement.RenderTranslatedStatement(tb);
        }

        internal override bool HasContent => _statement.HasContent;
    }

    internal sealed class CSharpCodeBuilderRawText(string statement) : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        private readonly string _statement = statement;
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            if (IndentationDepth > 0)
            {
                string indentationText = new string(' ', IndentationDepth * 4);
                tb.Append(indentationText);
            }
            tb.Append(_statement);
        }
    }

    internal sealed class CSharpClassBuilder() : CSharpCodeBuilder(TranslatedStatementKind.ClassDeclarationStatement)
    {
        private string? _className;
        private string? _baseClassName;
        private bool _public;
        private bool _sealed;

        public CSharpClassBuilder ClassName(string className)
        {
            _className = className;
            return this;
        }
        public CSharpClassBuilder BaseClassName(string baseClassName)
        {
            _baseClassName = baseClassName;
            return this;
        }
        public CSharpClassBuilder AsPublic()
        {
            _public = true;
            return this;
        }
        public CSharpClassBuilder AsSealed()
        {
            _sealed = true;
            return this;
        }
        internal CSharpClassBuilder AddProperty(CSharpCodeBuilder builder)
        {
            base.AddChildBuilder(builder);
            return this;
        }

        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            tb.Append(IndentationSpace);
            if (_public)
                tb.Append("public").Append(' ');
            if (_sealed)
                tb.Append("sealed").Append(' ');
            tb.Append("class").Append(' ').Append(_className);
            if (_baseClassName != null)
            {
                tb.Append(' ').Append(':').Append(' ').Append(_baseClassName);
            }
            tb.Append(NewLineNormalized);

            tb.Append(IndentationSpace).Append('{').Append(NewLineNormalized);

            // body
            foreach (CSharpCodeBuilder b in ChildBuilders())
            {
                b.RenderTranslatedStatement(tb);
                tb.Append(NewLineNormalized);
            }

            tb.Append(IndentationSpace).Append('}');//.Append(NewLineNormalized);
        }
    }
}
