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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

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
        public static string TranslateExecutable(CultureInfo culture, string scriptContent, IReadOnlyCollection<string> externalDependencies)
        {
            return TranslateCore(culture, scriptContent, externalDependencies, OuterScopeBlockTranslator.OutputTypeOptions.Executable, CommentsLogger(renderCommentsAboutUndeclaredVariables: true));
        }
        public static string TranslateWithoutScaffolding(CultureInfo culture, string scriptContent, NonNullImmutableList<string> externalDependencies)
        {
            return TranslateCore(culture, scriptContent, externalDependencies, OuterScopeBlockTranslator.OutputTypeOptions.WithoutScaffolding, CommentsLogger(renderCommentsAboutUndeclaredVariables: true));
        }
        internal static ILogInformation CommentsLogger(bool renderCommentsAboutUndeclaredVariables = true, ILogInformation? logger = null) => renderCommentsAboutUndeclaredVariables
            ? new CSharpCommentMakingLogger(logger ?? new ConsoleLogger())
            : new NullLogger();

        /// <summary>
        /// This Translate signature is what the others call into - it doesn't try to hide the fact that externalDependencies should be a NonNullImmutableList
        /// of strings and it requires an ILogInformation implementation to deal with logging warnings
        /// </summary>
        internal static string TranslateCore(
            CultureInfo culture,
            string scriptContent,
            IReadOnlyCollection<string> externalDependencies,
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

            CSharpName startNamespace = new CSharpName("TranslatedProgram");
            CSharpName startClassName = new CSharpName("Runner");
            CSharpName startMethodName = new CSharpName("Go");
            CSharpName runtimeDateLiteralValidatorClassName = new CSharpName("RuntimeDateLiteralValidator");
            CSharpName supportRefName = new CSharpName("_");
            CSharpName envClassName = new CSharpName("EnvironmentReferences");
            CSharpName envRefName = new CSharpName("_env");
            CSharpName outerClassName = new CSharpName("GlobalReferences");
            CSharpName outerRefName = new CSharpName("_outer");
            VBScriptNameRewriter nameRewriter = new DefaultVBScriptNameRewriter();
            TempValueNameGenerator tempNameGenerator = new DefaultTempValueNameGenerator().GenerateTempValueName;
            StatementTranslator statementTranslator = new StatementTranslator(supportRefName, envRefName, outerRefName, nameRewriter, tempNameGenerator, logger);
            OuterScopeBlockTranslator codeBlockTranslator = new OuterScopeBlockTranslator(
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
                externalDependencies.Select(name => new NameToken(false, name.ToUpperX(), 1)).ToNonNullImmutableList(), //1:First line
                outputType,
                logger
            );

            IReadOnlyList<ICodeBlock> parsedBlocks = Parse(culture, scriptContent);
            CSharpOutermostCodeBuilder outerBuilder = codeBlockTranslator.Translate(parsedBlocks);
            return outerBuilder.RenderTranslatedProgramCode();
        }

        /// <summary>
        /// This will return just the parsed VBScript content, it will not attempt any translation. It will never return null nor a set containing
        /// any null references. This may be used to analyse the structure of a script, if so desired.
        /// </summary>
        public static IReadOnlyList<ICodeBlock> Parse(CultureInfo culture, string scriptContent)
        {
            // Translate these tokens into ICodeBlock implementations (representing code VBScript structures)
            return CodeBlockHandler.RootBlock.Process(
                GetTokens(culture, scriptContent).ToList(),
                out string[]? _
            );
        }

        /// <summary>
        /// This will wrap log messages in C# comments (ensuring that there is no closing-comment symbol in the content which would invalidate the
        /// output as a comment). If a ConsoleLogger is used and the translated program content is sent to the console then this allows all of the
        /// output to be copy-pasted into a C# file for testing. Pretty rough and ready but can make things a little easier!
        /// </summary>
        private sealed class CSharpCommentMakingLogger : ILogInformation
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
            List<IToken> tokens = StringBreaker.SegmentString(culture, scriptContent);

            // Break down further into String, Comment, Atom and AbstractEndOfStatement tokens
            List<IToken> atomTokens = new List<IToken>();
            foreach (IToken? token in tokens)
            {
                if (token is UnprocessedContentToken unProccessed)
                {
                    atomTokens.AddRange(TokenBreaker.BreakUnprocessedToken(unProccessed));
                }
                else
                {
                    atomTokens.Add(token);
                }
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
        private readonly Dictionary<string, (RewriteEntry, CSharpName)> _entries = new Dictionary<string, (RewriteEntry, CSharpName)>();
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
            return RewriteNameX(value, line).Item1.RewrittenName;
        }
        private (RewriteEntry, CSharpName) RewriteNameX(string value, int line)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

#pragma warning disable CA1308 // Specify CultureInfo
            string key = value.ToLower(CultureInfo.InvariantCulture);
#pragma warning restore CA1308 // Specify CultureInfo
            if (_entries.TryGetValue(key, out (RewriteEntry, CSharpName) entry))
            {
                // already registered.
            }
            else
            {
                string rewrittenName = DefaultRuntimeSupportClassFactory.RewriteName(value);
                RewriteEntry entryR = new RewriteEntry(originalName: value, rewrittenName: rewrittenName, 0);
                entry = (entryR, new CSharpName(entryR.RewrittenName));
                _entries.Add(key, entry);
            }
            return entry;
        }
        public override CSharpName RewriteVBScriptName(NameToken name)
        {
            return RewriteNameX(name.Content, name.LineIndex).Item2;
        }
    }
}
