using System;
using System.Collections.Generic;
using System.Reflection;
using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.RuntimeSupport;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// A whole script that was parsed, translated and compiled once (see issue #187 and CompiledScriptCache). It knows its procedures without a
    /// further parse and every run gets a fresh compat layer and fresh environment / global references, so runs never share state - only the
    /// compiled code is shared (and a compiled script may be used by any number of parallel runs).
    /// </summary>
    internal sealed class CompiledScript
    {
        private readonly CodeLoadContext _codeLoadContext;

        private CompiledScript(string key, CompiledScriptSource source, CodeLoadContext codeLoadContext, string csCode, IReadOnlyDictionary<int, int> sourceLineMap, ScriptControlProcedures procedures, IReadOnlyDictionary<string, int> undeclaredVariableNames)
        {
            Key = key;
            ScriptContent = source.ScriptContent;
            _codeLoadContext = codeLoadContext;
            ProgramAssembly = codeLoadContext.LoadedAssembly;
            CsCode = csCode;
            SourceLineMap = sourceLineMap;
            Procedures = procedures;
            UndeclaredVariableNames = undeclaredVariableNames;
        }

        /// <summary>
        /// The cache key: a hash of the script text, the culture, the host object names (and members) and the translator / compiler options
        /// </summary>
        public string Key { get; }

        public string ScriptContent { get; }

        /// <summary>
        /// The Subs and Functions declared in the script (from the parse of the compilation - no further parse)
        /// </summary>
        public ScriptControlProcedures Procedures { get; }

        /// <summary>
        /// The loaded assembly of the compiled program (the same instance for every run of this compiled script)
        /// </summary>
        public Assembly ProgramAssembly { get; }

        public string CsCode { get; }
        public IReadOnlyDictionary<int, int> SourceLineMap { get; }
        public IReadOnlyDictionary<string, int> UndeclaredVariableNames { get; }

        /// <summary>
        /// Start a run with the compat layer of the caller (which the returned run disposes): fresh environment references with the host
        /// objects, then the outer scope of the script (its global statements) runs. Procedures are then called on the returned run.
        /// </summary>
        public CompiledScriptRun Run(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, IEnumerable<KeyValuePair<string, object>> externalObjects)
        {
            try
            {
                Type runnerType = ProgramAssembly.GetType("TranslatedProgram.Runner", true)!;
                RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(runnerType, compatLayer);
                EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();
                foreach (KeyValuePair<string, object> externalObject in externalObjects)
                    environmentReferences.InitializeExternalReference(externalObject.Key, externalObject.Value);
                GlobalReferencesBase globalReferences = runner.Run(environmentReferences);
                return new CompiledScriptRun(compatLayer, globalReferences);
            }
            catch
            {
                compatLayer.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Release the program: unloaded if its code load context can unload (issue #190), otherwise it stays loaded and only the reference goes
        /// </summary>
        internal void Unload() => _codeLoadContext.Unload();

        /// <summary>
        /// Parse, translate and compile the whole script (the parse of the translation also gives the procedures)
        /// </summary>
        internal static CompiledScript Compile(string key, CompiledScriptSource source, IHostCodeLoadContextHostService? codeLoadContextService = null)
        {
            ScriptControlConfiguration config = source.Configuration;
            NonNullImmutableList<string> externalDependencies = source.ExternalObjectNames.ToNonNullImmutableList();
            List<ExternalMemberMethodInfo> externalMemberMethods = new List<ExternalMemberMethodInfo>();
            foreach (KeyValuePair<string, string[]> objectMembers in source.ExternalObjectMembers)
            {
                foreach (string memberName in objectMembers.Value)
                    externalMemberMethods.Add(new ExternalMemberMethodInfo(objectMembers.Key, memberName));
            }

            // Translation warnings go to the trace listeners, as for ScriptControlClass
            var warningLogger = DefaultTranslator.CommentsLogger(true, new DefaultTranslator.DelegateWrappingWarningLogger(warningMessageText =>
            {
                System.Diagnostics.Trace.WriteLine(warningMessageText, "Skrypton");
            }));

            string programName = "CompiledScript_" + key.Substring(0, 16);
            if (config.TempEnabled)
                config.TempFileWriteAllText("CompiledScript", programName + ".vbs", source.ScriptContent);

            ScriptControlClass.ScriptTranslationIssuesCollector translationIssues = new ScriptControlClass.ScriptTranslationIssuesCollector();
            ScriptControlProcedures? procedures = null;
            string csCode = DefaultTranslator.TranslateCoreWithSourceLineMap(
                source.Culture,
                source.ScriptContent,
                externalDependencies,
                externalMemberMethods,
                OuterScopeBlockTranslator.OutputTypeOptions.Executable,
                DefaultTranslator.CreateTranslatorOptions(config._translationSuppression, translationIssues, config.UndeclaredVariablesAreErrors),
                warningLogger,
                out IReadOnlyDictionary<int, int> sourceLineMap,
                parsedBlocks => procedures = ScriptControlProcedures.FromBlocks(parsedBlocks));
            if (config.TempEnabled)
                config.TempFileWriteAllText("CompiledScript", programName + ".cs", csCode);

            CodeLoadContext asmctx = RoslynScriptControl.CompileCSharpProgram(config, programName, codeNumber: ScriptControlClass.NextExecNumber(), csCode, config.NoWarn, codeLoadContextService)!;
            return new CompiledScript(key, source, asmctx, csCode, sourceLineMap, procedures!, new Dictionary<string, int>(translationIssues._undeclaredNames, StringComparer.OrdinalIgnoreCase));
        }
    }
}
