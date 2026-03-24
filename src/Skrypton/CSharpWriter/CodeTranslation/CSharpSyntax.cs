using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    internal static class CSharpSyntaxFactory
    {
        internal static CSharpStatementBuilderConstructor CreateConstructor(int indentationDepth, int lineIndexOfStatementStartInSource) => CSharpCodeBuilder.Init(new CSharpStatementBuilderConstructor(), indentationDepth, lineIndexOfStatementStartInSource);

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
        internal static TBuilder CreateInitSetup<TBuilder>(TBuilder builder, int indentationDepth, int lineIndexOfStatementStartInSource, Action<TBuilder> setup) where TBuilder : CSharpCodeBuilder
        {
            Init(builder, indentationDepth, lineIndexOfStatementStartInSource);
            setup(builder);
            return builder;
        }
        protected void AddChildBuilder(CSharpCodeBuilder builder)
        {
            _builders.Add(builder);
        }
        protected IReadOnlyList<CSharpCodeBuilder> ChildBuilders() => _builders;

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

    internal sealed class CSharpStatementBuilderConstructor() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
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

    internal abstract class CSharpBlockCodeBuilderT<TBuilder>() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        protected abstract TBuilder That { get; }
        protected TBuilder AddChildBuilderX(CSharpCodeBuilder builder)
        {
            base.AddChildBuilder(builder);
            return That;
        }
        protected static TB AddChildBuilderT<TB>(TB parent, CSharpCodeBuilder builder) where TB : CSharpBlockCodeBuilderT<TBuilder>
        {
            parent.AddChildBuilder(builder);
            return parent;
        }
        internal TBuilder AddBuilder(CSharpCodeBuilder builder)
        {
            return AddChildBuilderX(builder);
        }
        public TBuilder AddRange(IReadOnlyCollection<TranslatedStatement> values)
        {
            foreach (TranslatedStatement value in values)
            {
                AddChildBuilder(new CSharpCodeBuilderWrap(value));
            }
            return That;
        }
        public TBuilder Add(TranslatedStatement value)
        {
            return AddChildBuilderX(new CSharpCodeBuilderWrap(value));
        }
        internal TBuilder AddRawText(string rawText, int indentationDepth, int lineIndexOfStatementStartInSource)
        {
            return AddChildBuilderX(CSharpSyntaxFactory.FromRawText(rawText, indentationDepth, lineIndexOfStatementStartInSource));
        }

        internal CSharpClassBuilder CreateClass(int indentationDepth, int lineIndexOfStatementStartInSource) => CSharpCodeBuilder.Init(new CSharpClassBuilder(), indentationDepth, lineIndexOfStatementStartInSource > 0 ? lineIndexOfStatementStartInSource : LineIndexOfStatementStartInSource);

        internal TBuilder AddAssignmentStatement(int indentationDepth, int lineIndexOfStatementStartInSource, Action<CSharpAssignmentStatement> setup) => AddChildBuilderX(CSharpCodeBuilder.CreateInitSetup(new CSharpAssignmentStatement(), indentationDepth, lineIndexOfStatementStartInSource, setup));
        internal TBuilder AddVariableDeclaration(int indentationDepth, int lineIndexOfStatementStartInSource, Action<CSharpVariableDeclarationStatement> setup) => AddChildBuilderX(CSharpCodeBuilder.CreateInitSetup(new CSharpVariableDeclarationStatement(), indentationDepth, lineIndexOfStatementStartInSource, setup));
        internal TBuilder AddMethodInvocationStatement(int indentationDepth, int lineIndexOfStatementStartInSource, Action<CSharpInvocationStatement> setup) => AddChildBuilderX(CSharpCodeBuilder.CreateInitSetup(new CSharpInvocationStatement(), indentationDepth, lineIndexOfStatementStartInSource, setup));

        private List<CSharpCodeBuilder> RemoveRunsOfBlankLines()
        {
            List<CSharpCodeBuilder> children = ChildBuilders().ToList();

            return children
                .Select((s, i) => ((i == 0) || (s.HasContent) || (children[i - 1].HasContent)) ? s : null)
                .Where(s => s != null)
                .Select(s => s!)
                .ToList();
        }

        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            foreach (CSharpCodeBuilder s in ChildBuilders())
            {
                s.RenderTranslatedStatement(tb);
                tb.Append(NewLineNormalized);
            }
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

    internal abstract class CSharpOutermostCodeBuilder() : CSharpBlockCodeBuilderT<CSharpOutermostCodeBuilder>
    {
        protected override CSharpOutermostCodeBuilder That =>  this;
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            throw new NotImplementedException();
        }
    }
    internal sealed class CSharpProgramCodeBuilder : CSharpOutermostCodeBuilder
    {
        internal CSharpProgramCodeBuilder AddUsing<T>() => AddChildBuilderT(this, CSharpSyntaxFactory.FromRawText($"using {typeof(T).Namespace!};", 0, 0));
    }
    internal sealed class CSharpScaffoldingCodeBuilder : CSharpOutermostCodeBuilder
    {

    }

    internal sealed class CSharpVariableDeclarationStatement() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        private string? _name;
        private string? _typeText;
        private string? _initializationText;
        internal CSharpVariableDeclarationStatement VariableName(string name)
        {
            _name = name;
            return this;
        }

        internal CSharpVariableDeclarationStatement VariableType<T>()
        {
            if (typeof(T) == typeof(int))
            {
                _typeText = "int";
            }
            else
            {
                _typeText = typeof(T).FullName;
            }
            return this;
        }

        internal CSharpVariableDeclarationStatement VariableInitialization(string initializationText)
        {
            _initializationText = initializationText;
            return this;
        }
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            tb.Append(IndentationSpace).Append(_typeText).Append(' ').Append(_name);
            if (!string.IsNullOrEmpty(_initializationText))
                tb.Append(' ').Append('=').Append(' ').Append(_initializationText);
        }

    }
    internal sealed class CSharpAssignmentStatement() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        private string? _referenceName;
        public CSharpAssignmentStatement ReferenceName(string referenceName)
        {
            _referenceName = referenceName;
            return this;
        }
        private string? _expressionText;
        public CSharpAssignmentStatement ExpressionText(string expressionText)
        {
            _expressionText = expressionText;
            return this;
        }
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            // $"{_supportRefName.Name}.RELEASEERRORTRAPPINGTOKEN({scopeAccessInformation.ErrorRegistrationTokenIfAny.Name});
            tb.Append(IndentationSpace).Append(_referenceName).Append(' ').Append('=').Append(' ').Append(_expressionText).Append(';');
            throw new NotImplementedException();
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

    internal sealed class CSharpPropertyDeclarationBuilder() : CSharpCodeBuilder(TranslatedStatementKind.PropertyDeclarationStatement)
    {
        private string? _propertyName;
        public CSharpPropertyDeclarationBuilder PropertyName(string propertyName)
        {
            _propertyName = propertyName;
            return this;
        }
        private bool _public;
        public CSharpPropertyDeclarationBuilder AsPublic()
        {
            _public = true;
            return this;
        }
        private string? _propertyTypeName;
        public CSharpPropertyDeclarationBuilder PropertyTypeName(string propertyTypeName)
        {
            _propertyTypeName = propertyTypeName;
            return this;
        }

        private string? _getterText;
        public CSharpPropertyDeclarationBuilder PublicGetter(string getterText)
        {
            _getterText = getterText;
            return this;
        }

        private string? _setterText;
        public CSharpPropertyDeclarationBuilder InternalSetter(string setterText)
        {
            _setterText = setterText;
            return this;
        }
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            // public object " + v.RewrittenName + " { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
            tb.Append(IndentationSpace);
            if (_public)
                tb.Append("public").Append(' ');
            tb.Append(_propertyTypeName).Append(' ').Append(_propertyName).Append(' ').Append('{');
            if (!string.IsNullOrEmpty(_getterText))
                tb.Append(' ').Append("get").Append(' ').Append("=>").Append(' ').Append(_getterText).Append(';');
            if (!string.IsNullOrEmpty(_setterText))
                tb.Append(' ').Append("internal").Append(' ').Append("set").Append(' ').Append("=>").Append(' ').Append(_setterText).Append(';');
            tb.Append(' ').Append('}');
        }
    }

    internal sealed class CSharpClassBuilder() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
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
        internal CSharpClassBuilder AddProperty(int line, Action<CSharpPropertyDeclarationBuilder> setup)
        {
            CSharpPropertyDeclarationBuilder builder = CSharpPropertyDeclarationBuilder.Init(new CSharpPropertyDeclarationBuilder(), IndentationDepth + 1, line);
            setup(builder);
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

    internal sealed class CSharpInvocationStatement() : CSharpCodeBuilder(TranslatedStatementKind.RawText)
    {
        private string? _targetName;
        private string? _methodName;
        public CSharpInvocationStatement TargetName(string targetName)
        {
            _targetName = targetName;
            return this;
        }
        public CSharpInvocationStatement MethodName(string methodName)
        {
            _methodName = methodName;
            return this;
        }
        public CSharpInvocationStatement AddParameterVariableReference(string name)
        {
            AddChildBuilder(CSharpSyntaxFactory.FromRawText(name, 0, LineIndexOfStatementStartInSource));
            return this;
        }
        public CSharpInvocationStatement AddParameters<T>(T[] sources, Func<T, string> gen)
        {
            foreach (var source in sources)
            {
                string text = gen(source);
                AddChildBuilder(CSharpSyntaxFactory.FromRawText(text, 0, LineIndexOfStatementStartInSource));
            }
            return this;
        }
        protected override void DoBuildTranslatedStatement(StringBuilder tb)
        {
            // _.RELEASEERRORTRAPPINGTOKEN(errOn);
            tb.Append(IndentationSpace).Append(_targetName).Append('.').Append(_methodName);
            tb.Append('(');
            var parameters = ChildBuilders();
            for (int ixPrm = 0; ixPrm < parameters.Count; ixPrm++)
            {
                var prm = parameters[ixPrm];
                if (ixPrm > 0)
                    tb.Append(", ");
                prm.RenderTranslatedStatement(tb);
            }
            tb.Append(')').Append(';');
        }
    }
}