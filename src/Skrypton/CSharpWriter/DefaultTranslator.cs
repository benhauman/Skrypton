using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.CodeTranslation.StatementTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.CSharpWriter.Logging;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.LegacyParser.ContentBreaking;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using Skrypton.RuntimeSupport;
using Skrypton.StageTwoParser.TokenCombining.NumberRebuilding;
using Skrypton.StageTwoParser.TokenCombining.OperatorCombinations;

namespace Skrypton.CSharpWriter
{
    public sealed class DefaultTranslator
    {
        /// <summary>
        /// This will attempt to translate VBScript content into C# using the default configurations, probably the best place to start (it uses the
        /// DefaultRuntimeSupportClassFactory for name rewriting, so that same name rewriter must be used to execute the output generated here). If
        /// there are any runtime references that are known to be present (such as WScript when run within CScript at the command line, or Request,
        /// Response, Session, etc.. when run within ASP) then specify their names in the externalDependencies set - this will prevent warnings
        /// being logged in relation to the absence of their definition in the source.
        /// </summary>
        public static NonNullImmutableList<TranslatedStatement> TranslateExecutable(CultureInfo culture, string scriptContent, NonNullImmutableList<string> externalDependencies)
        {
            return TranslateCore(culture, scriptContent, externalDependencies, OuterScopeBlockTranslator.OutputTypeOptions.Executable, CommentsLogger(renderCommentsAboutUndeclaredVariables: true));
        }
        public static NonNullImmutableList<TranslatedStatement> TranslateWithoutScaffolding(CultureInfo culture, string scriptContent, NonNullImmutableList<string> externalDependencies)
        {
            return TranslateCore(culture, scriptContent, externalDependencies, OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding, CommentsLogger(renderCommentsAboutUndeclaredVariables: true));
        }
        internal static ILogInformation CommentsLogger(bool renderCommentsAboutUndeclaredVariables = true, ILogInformation logger = null) => renderCommentsAboutUndeclaredVariables
            ? new CSharpCommentMakingLogger(logger ?? new ConsoleLogger())
            : new NullLogger();

        ///// <summary>
        ///// This Translate signature exists to provide an extremely simple way to get code translated - it is used in some of the examples so that
        ///// there's a way to get to translating before worrying about what the NonNullImmutableList type is all about
        ///// </summary>
        //internal static NonNullImmutableList<TranslatedStatement> TranslateX(CultureInfo culture, string scriptContent, string[] externalDependencies,
        //    OuterScopeBlockTranslator.OutputTypeOptions outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable,
        //    bool renderCommentsAboutUndeclaredVariables = true)
        //{
        //    if (externalDependencies == null)
        //        throw new ArgumentNullException("externalDependencies");

        //    return Translate(culture, scriptContent, externalDependencies.ToNonNullImmutableList(), outputType, true);
        //}

        /// <summary>
        /// This Translate signature exists to provide a slightly-simpler way to specify a custom warning logger (by providing a simple delegate,
        /// rather than having to provide an ILogInformation implementation)
        /// </summary>
        //lubo:public static NonNullImmutableList<TranslatedStatement> Translate(
        //lubo:    CultureInfo culture,
        //lubo:    string scriptContent,
        //lubo:    string[] externalDependencies,
        //lubo:    Action<string> warningLogger,
        //lubo:    OuterScopeBlockTranslator.OutputTypeOptions outputType = OuterScopeBlockTranslator.OutputTypeOptions.Executable)
        //lubo:{
        //lubo:    if (externalDependencies == null)
        //lubo:        throw new ArgumentNullException("externalDependencies");
        //lubo:    if (warningLogger == null)
        //lubo:        throw new ArgumentNullException("warningLogger");
        //lubo:
        //lubo:    return Translate(culture, scriptContent, externalDependencies.ToNonNullImmutableList(), outputType, new DelegateWrappingWarningLogger(warningLogger));
        //lubo:}

        /// <summary>
        /// This Translate signature is what the others call into - it doesn't try to hide the fact that externalDependencies should be a NonNullImmutableList
        /// of strings and it requires an ILogInformation implementation to deal with logging warnings
        /// </summary>
        internal static NonNullImmutableList<TranslatedStatement> TranslateCore(
            CultureInfo culture,
            string scriptContent,
            NonNullImmutableList<string> externalDependencies,
            OuterScopeBlockTranslator.OutputTypeOptions outputType,
            ILogInformation logger
            )
        {
            if (scriptContent == null)
                throw new ArgumentNullException(nameof(scriptContent));
            if (externalDependencies == null)
                throw new ArgumentNullException(nameof(externalDependencies));
            if ((outputType != OuterScopeBlockTranslator.OutputTypeOptions.Executable) && (outputType != OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding))
                throw new ArgumentOutOfRangeException(nameof(outputType));
            if (logger == null)
                throw new ArgumentNullException(nameof(logger));

            var startNamespace = new CSharpName("TranslatedProgram");
            var startClassName = new CSharpName("Runner");
            var startMethodName = new CSharpName("Go");
            var runtimeDateLiteralValidatorClassName = new CSharpName("RuntimeDateLiteralValidator");
            var supportRefName = new CSharpName("_");
            var envClassName = new CSharpName("EnvironmentReferences");
            var envRefName = new CSharpName("_env");
            var outerClassName = new CSharpName("GlobalReferences");
            var outerRefName = new CSharpName("_outer");
            VBScriptNameRewriter nameRewriter = new DefaultVBScriptNameRewriter();
            TempValueNameGenerator tempNameGenerator = new DefaultTempValueNameGenerator().GenerateTempValueName;
            var statementTranslator = new StatementTranslator(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, logger);
            var codeBlockTranslator = new OuterScopeBlockTranslator(
                startNamespace,
                startClassName,
                startMethodName,
                runtimeDateLiteralValidatorClassName,
                supportRefName,
                envClassName,
                envRefName,
                outerClassName,
                outerRefName,
                nameRewriter,
                tempNameGenerator,
                statementTranslator,
                new ValueSettingStatementsTranslator(supportRefName, envRefName, outerRefName, nameRewriter, statementTranslator, logger),
                externalDependencies.Select(name => new NameToken(false, name.ToUpperX(), 0)).ToNonNullImmutableList(),
                outputType,
                logger
            );

            return codeBlockTranslator.Translate(
                Parse(culture, scriptContent).ToNonNullImmutableList()
            );
        }

        /// <summary>
        /// This will return just the parsed VBScript content, it will not attempt any translation. It will never return null nor a set containing
        /// any null references. This may be used to analyse the structure of a script, if so desired.
        /// </summary>
        public static IEnumerable<ICodeBlock> Parse(CultureInfo culture, string scriptContent)
        {
            // Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
            string[] endSequenceMet;
            return CodeBlockHandler.RootBlock.Process(
                GetTokens(culture, scriptContent).ToList(),
                out endSequenceMet
            );
        }

        /// <summary>
        /// This will wrap log messages in C# comments (ensuring that there is no closing-comment symbol in the content which would invalidate the
        /// output as a comment). If a ConsoleLogger is used and the translated program content is sent to the console then this allows all of the
        /// output to be copy-pasted into a C# file for testing. Pretty rough and ready but can make things a little easier!
        /// </summary>
        private class CSharpCommentMakingLogger : ILogInformation
        {
            private readonly ILogInformation _logger;
            public CSharpCommentMakingLogger(ILogInformation logger)
            {
                _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            }
            public void Warning(string content)
            {
                if (!string.IsNullOrWhiteSpace(content))
                    content = "/* " + content.Replace("*/", "*") + " */";
                _logger.Warning(content);
            }
        }

        private static IToken[] GetTokens(CultureInfo culture, string scriptContent)
        {
            // Break down content into String, Comment and UnprocessedContent tokens
            var tokens = StringBreaker.SegmentString(culture, scriptContent);

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            var atomTokens = new List<IToken>();
            foreach (var token in tokens)
            {
                if (token is UnprocessedContentToken)
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken((UnprocessedContentToken)token));
                else
                    atomTokens.Add(token);
            }

            return NumberRebuilder.Rebuild(OperatorCombiner.Combine(atomTokens)).ToArray();
        }

        internal sealed class DelegateWrappingWarningLogger : ILogInformation
        {
            private readonly Action<string> _warningLogger;
            public DelegateWrappingWarningLogger(Action<string> warningLogger)
            {
                _warningLogger = warningLogger ?? throw new ArgumentNullException(nameof(warningLogger));
            }

            public void Warning(string content)
            {
                _warningLogger(content);
            }
        }
    }

    internal sealed class DefaultTempValueNameGenerator
    {
        private readonly Dictionary<string, string> _names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //private int _tempNameGeneratorNextNumber;
        public DefaultTempValueNameGenerator()
        {
        }
        public CSharpName GenerateTempValueName(CSharpName optionalPrefix, ScopeAccessInformation scopeAccessInformation)
        {
            // To get unique names for any given translation, a running counter is maintained and appended to the end of the generated
            // name. This is only run during translation (this code is not used during execution) so there will be a finite number of
            // times that this is called (so there should be no need to worry about the int value overflowing!)
            //_tempNameGeneratorNextNumber++;
            //string numberSuffix = (_tempNameGeneratorNextNumber == 1) ? "" : _tempNameGeneratorNextNumber.ToString(CultureInfo.InvariantCulture);
            //string name = ((optionalPrefix == null) ? "temp" : optionalPrefix.Name) + numberSuffix;
            //return new CSharpName(name);
            int tempNameGeneratorNextNumber = 1;
            string prefix = optionalPrefix?.Name ?? "temp";
            string name = prefix;
            while (_names.ContainsKey(name))
            {
                tempNameGeneratorNextNumber++;
                string numberSuffix = (tempNameGeneratorNextNumber == 1) ? "" : tempNameGeneratorNextNumber.ToString(CultureInfo.InvariantCulture);
                name = prefix + numberSuffix;
            }
            _names.Add(name, prefix);
            return new CSharpName(name);
        }
    }

    public sealed class DefaultVBScriptNameRewriter : VBScriptNameRewriter
    {
        private readonly Dictionary<string, RewriteEntry> _entries = new Dictionary<string, RewriteEntry>();
        private sealed class RewriteEntry
        {
            public string OriginalName { get; }
            public string RewrittenName { get; }
            public int LineIndex { get; }

            public RewriteEntry(string originalName, string rewrittenName, int lineIndex)
            {
                OriginalName = originalName ?? throw new ArgumentNullException(nameof(originalName));
                RewrittenName = rewrittenName ?? throw new ArgumentNullException(nameof(rewrittenName));
                LineIndex = lineIndex;
            }
        }
        public DefaultVBScriptNameRewriter()
        {
        }
        internal string RewriteName(string value, int line)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            string key = value.ToLower();
            if (_entries.TryGetValue(key, out RewriteEntry entry))
            {
                // already registered.
            }
            else
            {
                string rewrittenName = DefaultRuntimeSupportClassFactory.RewriteName(value);
                entry = new RewriteEntry(originalName: value, rewrittenName: rewrittenName, 0);
                _entries.Add(key, entry);
            }
            return entry.RewrittenName;
        }
        public override CSharpName RewriteVBScriptName(NameToken name)
        {
            return new CSharpName(RewriteName(name.Content, name.LineIndex));
        }
    }
}
