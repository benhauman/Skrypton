using Skrypton.CSharpWriter;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using Skrypton.CSharpWriter.CodeTranslation.BlockTranslators;

namespace Skrypton.ScriptControlSupport
{
    public sealed class ScriptControlClass : IScriptControl
    {
        private int _timeout = -1; // -1 means no timeout (infinite execution time)
        private bool _allowUI;// Disable UI by default for security reasons (MessageBox, InputBox, etc.)
        private string _language = "";// ('VBScript') No scripting language is selected by default. It must be set explicitly, otherwise script execution will fail
        private ScriptControlStates _state = ScriptControlStates.Initialized;
        private int _sitehWnd; // required (Must be a valid HWND (window handle)) if allowUI is true otherwise ignored. No host window is associated by default => 0
        private bool _useSafeSubset = true; // 'true': Script runs in safe modeUse safe subset of the scripting language (if supported). 'false': Use full language features. Safe subset is used by default. Potentially dangerous objects and operations are blocked

        //private Error _error; // Default value: null (or Nothing in VBScript) — no error has occurred yet.
        //#pragma warning disable CS0414 // The field is assigned but its value is never used
        //        private object _codeObject = null; // Default value: null (or Nothing in VBScript) — no code object has been set yet. This allows you to interact with script members directly, instead of using Run or ExecuteStatement.
        //#pragma warning restore CS0414 // The field is assigned but its value is never used
        private readonly ScriptControlConfiguration _config;

        public ScriptControlClass(IRuntimeHost engineRuntimeHost, IRuntimeLogger engineRuntimeLogger, CultureInfo engineCulture, ScriptControlConfiguration config)
        {
            EngineRuntimeHost = engineRuntimeHost ?? throw new ArgumentNullException(nameof(engineRuntimeHost));
            EngineCulture = engineCulture ?? throw new ArgumentNullException(nameof(engineCulture));
            EngineRuntimeLogger = engineRuntimeLogger ?? throw new ArgumentNullException(nameof(engineRuntimeLogger));
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        string IScriptControl.Language { get => _language; set => _language = value; }
        ScriptControlStates IScriptControl.State
        {
            get => _state;

            // Starting or stopping the script engine manually
            // Setting State = ScriptControlStates.Initialized can reset the engine, clearing global variables and functions.
            set => throw new NotImplementedException();
        }
        int IScriptControl.SitehWnd { get => _sitehWnd; set => _sitehWnd = value; }
        int IScriptControl.Timeout { get => _timeout; set => _timeout = value; }
        bool IScriptControl.AllowUI { get => _allowUI; set => _allowUI = value; }
        bool IScriptControl.UseSafeSubset { get => _useSafeSubset; set => throw new NotImplementedException(); }

        // The list of all script modules loaded into the control
        // When you use AddCode(...), the code is added to the default module automatically.
        // All functions, subs, and variables you define live there.
        // Everything goes into the default module unless you load a script file that has modules defined.
        // No named modules (by default).
        // VBScript does not really support named modules like VBA, so everything still ends up in the default module in ScriptControl.
        // If you want true separate modules, you would need multiple ScriptControl instances.
        // !!! You cannot have two modules with the same function name and signature that behave differently
        //    => 1. Use separate ScriptControl instances
        //    => 2. Use different function names
        //    => 3. Wrap in classes/objects (VBScript only) and then instantiate them in vbs code.
        // The Modules collection is mostly for inspection; it does not separate namespaces.
        // VBScript in ScriptControl does not support namespaces or module scoping like C#.
        //
        //   sc.AddCode(@"
        //   Module MathModule
        //       Function Multiply(x, y)
        //           Multiply = x * y
        //       End Function
        //   End Module
        //   ");
        //
        Modules IScriptControl.Modules => throw new NotImplementedException();

        private readonly ScriptControlError _error = new ScriptControlError();
        private IReadOnlyDictionary<int, int>? _lastSourceLineMap;
        private string? _lastScriptContent;

        /// <summary>
        /// The error raised by the most recent ExecuteStatement call (Number is zero if there was none), see issue #47
        /// </summary>
        Error? IScriptControl.Error => _error;

        object IScriptControl.CodeObject => throw new NotImplementedException();

        /// <summary>
        /// The Subs and Functions declared in the code added with AddCode (see issue #47)
        /// </summary>
        Procedures IScriptControl.Procedures => ScriptControlProcedures.FromScript(EngineCulture, _code.ToString());

        void IScriptControl._AboutBox()
        {
            throw new NotImplementedException();
        }

        private readonly Dictionary<string, object> _addedObjects = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string[]> _addedObjectMembers = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        void IScriptControl.AddObject(string objectName, object objectInstance, bool addMembers)
        {
            if (objectName == null) throw new ArgumentNullException(nameof(objectName));
            if (string.IsNullOrEmpty(objectName)) throw new ArgumentException("Value cannot be null or empty.", nameof(objectName));

            if (_addedObjects.ContainsKey(objectName))
            {
                // “An object with this name already exists in the script namespace.”
                throw new InvalidOperationException($"An object with this name already exists in the script namespace. '{objectName}' HR:0x800A03EC (SCRIPT_E_DUPLICATEOBJECTNAME)"); // SCRIPT_E_DUPLICATEOBJECTNAME
            }
            if (addMembers)
            {
                string[] methods = objectInstance.GetType().GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly).Where(x => x.DeclaringType != typeof(object)).Select(x => x.Name).Distinct().ToArray();
                if (methods.Length > 0)
                {
                    _addedObjectMembers.Add(objectName, methods);
                }
            }
            _addedObjects[objectName] = objectInstance ?? throw new ArgumentNullException(nameof(objectInstance));

            if (addMembers)
            {
                //throw new NotImplementedException();
            }
        }

        void IScriptControl.Reset()
        {
            ResetCore();
        }

        private void ResetCore()
        {
            throw new NotImplementedException();
        }

        private readonly StringBuilder _code = new StringBuilder();
        void IScriptControl.AddCode(string code)
        {
            if (string.IsNullOrEmpty(code)) throw new ArgumentException("Value cannot be null or empty.", nameof(code));
            _code.Append(code);
        }

        object IScriptControl.Eval(string Expression)
        {
            throw new NotImplementedException();
        }

        public void TestTranslatedStatement(string programName, string csCode, string[] nowarn, bool doRun, Action<GlobalReferencesBase> testHandler)
        {
            if (nowarn == null) throw new ArgumentNullException(nameof(nowarn));
            if (testHandler == null) throw new ArgumentNullException(nameof(testHandler));
            if (string.IsNullOrEmpty(csCode)) throw new ArgumentException("Value cannot be null or empty.", nameof(csCode));
            UnloadableAssemblyLoadContextContext? asmctx = RoslynScriptControl.CompileCSharpProgram(_config, programName, codeNumber: _lastExecNumber, csCode, nowarn);
            try
            {
                using Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer = CreateCompatLayer();
                Type tRunner = asmctx!.LoadedAssembly.GetType("TranslatedProgram.Runner", true);
                RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(tRunner, compatLayer);

                EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();

                foreach (KeyValuePair<string, object> externalReferencesEntry in _addedObjects)
                {
                    environmentReferences.InitializeExternalReference(externalReferencesEntry.Key, externalReferencesEntry.Value);
                }

                if (doRun)
                {
                    GlobalReferencesBase globalRefs = runner.Run(environmentReferences);
                    testHandler(globalRefs);
                }
            }
            finally
            {
                asmctx?.UnloadContextCollectAndWait();
                asmctx = null;
            }
        }
        private static int _lastExecNumber;

        /// <summary>
        /// The assembly of the program that was run by the most recent ExecuteStatement call (programs are cached, see CompiledProgramCache)
        /// </summary>
        public Assembly? LastExecutedProgramAssembly { get; private set; }

        void IScriptControl.ExecuteStatement(string statement)
        {
            if (string.IsNullOrEmpty(statement))
                throw new ArgumentException("Value cannot be null or empty.", nameof(statement));

            // As with MSScriptControl, Error describes the error raised by the most recent call (and is cleared by a call that succeeds), see issue #47
            _error.Clear();
            Assembly? programAssembly = null;
            try
            {
                ExecuteStatementCore(statement, out programAssembly);
            }
            catch (Exception e)
            {
                _error.SetFrom(e, programAssembly, _lastSourceLineMap, _lastScriptContent);
                throw;
            }
        }

        private void ExecuteStatementCore(string statement, out Assembly? programAssemblyIfCompiled)
        {
            programAssemblyIfCompiled = null;
            //RoslynScriptControl sc = StartAsync(statement, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            //sc.ExecuteStatementAsync(statement)
            string csCode = GenerateCSharpCode(statement, _addedObjectMembers);
            //RoslynScriptControl sc = new RoslynScriptControl();
            //sc.ExecuteStatementAsync(csCode, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            // Identical programs are compiled and loaded once and then reused (see issue #43 and CompiledProgramCache) - each run gets its own compat
            // layer and its own environment / global references, so nothing is shared between runs other than the compiled code
            string cacheKey = CompiledProgramCache.GetKey(csCode, _config.NoWarn, _config.TempEnabled, _config.EnabledLoadFromDisk);
            Assembly programAssembly = CompiledProgramCache.GetOrAdd(cacheKey, () =>
            {
                int execNumber = Interlocked.Increment(ref _lastExecNumber); // threadsafe
                UnloadableAssemblyLoadContextContext? asmctx = RoslynScriptControl.CompileCSharpProgram(_config, "TempScriptProgram", codeNumber: execNumber, csCode, _config.NoWarn);
                return asmctx!.LoadedAssembly;
            });
            LastExecutedProgramAssembly = programAssembly;
            programAssemblyIfCompiled = programAssembly;

            //DefaultRuntimeSupportClassFactory defaultRuntimeSupportClassFactoryInstance = _defaultRuntimeSupportClassFactoryProvider(EngineRuntimeHost, EngineRuntimeLogger, EngineCulture);
            using Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer = CreateCompatLayer();// defaultRuntimeSupportClassFactoryInstance.RuntimeHost, defaultRuntimeSupportClassFactoryInstance.RuntimeLogger, defaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever, EngineCulture);
            Type tRunner = programAssembly.GetType("TranslatedProgram.Runner", true)!;
            RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(tRunner, compatLayer);

            EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();

            foreach (KeyValuePair<string, object> externalReferencesEntry in _addedObjects)
            {
                environmentReferences.InitializeExternalReference(externalReferencesEntry.Key, externalReferencesEntry.Value);
            }

            GlobalReferencesBase globalRefs = runner.Run(environmentReferences);
            // Note: There is no forced GC.Collect / WaitForPendingFinalizers here any more - the compiled program is loaded into the default load
            // context (so it could never be unloaded by a collection) and a blocking full GC per call is expensive on a busy server (see issue #43)

            //if (!weakRef.IsAlive)
            //    Console.WriteLine("ALC successfully unloaded");
            //else
            //    Console.WriteLine("ALC still alive");
        }

        private DefaultRuntimeFunctionalityProvider CreateCompatLayer()
        {
            DefaultRuntimeSupportClassFactory defaultRuntimeSupportClassFactoryInstance = DefaultRuntimeSupportClassFactory.Create(EngineRuntimeHost, EngineRuntimeLogger, EngineCulture);
            DefaultRuntimeFunctionalityProvider compatLayer = new DefaultRuntimeFunctionalityProvider(EngineRuntimeHost, EngineRuntimeLogger, defaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever, EngineCulture);
            if (_setupDefaultRuntimeFunctionalityProvider != null)
            {
                _setupDefaultRuntimeFunctionalityProvider.Invoke(compatLayer);
            }
            return compatLayer;
        }
        private Action<DefaultRuntimeFunctionalityProvider>? _setupDefaultRuntimeFunctionalityProvider;
        public void TestSetDefaultRuntimeFunctionalityProviderSetup(Action<DefaultRuntimeFunctionalityProvider> setupDefaultRuntimeFunctionalityProvider)
        {
            _setupDefaultRuntimeFunctionalityProvider = setupDefaultRuntimeFunctionalityProvider ?? throw new ArgumentNullException(nameof(setupDefaultRuntimeFunctionalityProvider));
        }
        object IScriptControl.Run(string procedureName, ref object[] parameters)
        {
            //RoslynScriptControl sc = StartAsync(null, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            //return RunProcedure(procedureName);
            throw new NotImplementedException();
        }

        public static object RunProcedure(GlobalReferencesBase globalRefs, string procedureName, object[] parameters)
        {
            if (globalRefs == null) throw new ArgumentNullException(nameof(globalRefs));
            if (parameters == null) throw new ArgumentNullException(nameof(parameters));

            var mi = globalRefs.GetMethodInfoByNameAndArgs(procedureName, parameters);
            return mi.Invoke(globalRefs, parameters);
        }

        public IRuntimeHost EngineRuntimeHost { get; private set; }
        public IRuntimeLogger EngineRuntimeLogger { get; set; }
        public CultureInfo EngineCulture { get; set; }
        /*private Task<RoslynScriptControl> StartAsync(string statementOrNull, CancellationToken cancellationToken)
        {
            string csCode = GenerateCSharpCode(statementOrNull);
            //RoslynScriptControl sc = new RoslynScriptControl();
            //sc.
            //sc.ExecuteStatementAsync(csCode, cancellationToken).ConfigureAwait(false).GetAwaiter().GetResult();
            return Task.FromResult(sc);
            //return Task.CompletedTask;
        }*/
        public string TestGenerateCSharpCode()
        {
            return GenerateCSharpCode(null, _addedObjectMembers);
        }

        int _translationCount;
        private string GenerateCSharpCode(string? statementOrNull, Dictionary<string, string[]> addedObjectMembers)
        {
            string scriptContent = _code.ToString(); // Assume this is populated with the script code to be parsed
            if (!string.IsNullOrEmpty(scriptContent))
            {
                if (statementOrNull != null)
                {
                    scriptContent += NewLineNormalized;
                }
            }
            if (statementOrNull != null)
            {
                scriptContent += statementOrNull;
            }
            NonNullImmutableList<string> externalDependencies = _addedObjects.Keys.ToArray().ToNonNullImmutableList();
            List<ExternalMemberMethodInfo> externalMemberMethods = new List<ExternalMemberMethodInfo>();
            foreach (var xx in addedObjectMembers)
            {
                foreach (string externalMemberName in xx.Value)
                {
                    externalMemberMethods.Add(new ExternalMemberMethodInfo(xx.Key, externalMemberName));
                }
            }

            var warningLogger = Skrypton.CSharpWriter.DefaultTranslator.CommentsLogger(true, new Skrypton.CSharpWriter.DefaultTranslator.DelegateWrappingWarningLogger(warningMessageText =>
            {
                Console.WriteLine(warningMessageText);
            }));

            _translationCount++;

            string programSourceFileNameNoExt = $"ProgramName{(_translationCount > 1 ? _translationCount.ToString(CultureInfo.InvariantCulture) : string.Empty)}";
            if (_config.TempEnabled)
            {
                _config.TempFileWriteAllText($"ProgramName", $"{programSourceFileNameNoExt}.vbs", scriptContent);
            }
            _translationIssues.Clear();
            string[] suppressions = _config._translationSuppression;
            _lastScriptContent = scriptContent;
            _lastSourceLineMap = null;
            string csCode = Skrypton.CSharpWriter.DefaultTranslator.TranslateCoreWithSourceLineMap(EngineCulture, scriptContent, externalDependencies, externalMemberMethods, CSharpWriter.CodeTranslation.BlockTranslators.OuterScopeBlockTranslator.OutputTypeOptions.Executable, DefaultTranslator.CreateTranslatorOptions(suppressions, _translationIssues), warningLogger, out IReadOnlyDictionary<int, int> sourceLineMap);
            _lastSourceLineMap = sourceLineMap;
            if (_config.TempEnabled)
            {
                _config.TempFileWriteAllText($"ProgramName", $"{programSourceFileNameNoExt}.cs", csCode);
            }
            return csCode;
        }
        private const char NewLineNormalized = '\n';
        private sealed class ScriptTranslationIssuesCollector : TranslationIssuesCollectorBase
        {
            internal Dictionary<string, int> _undeclaredNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public override void UndeclaredNamedReferenceDetected(string errorKey, string referenceName, int lineIndex)
            {
                if (string.IsNullOrEmpty(referenceName)) throw new ArgumentException("Value cannot be null or empty.", nameof(referenceName));
                if (!_undeclaredNames.ContainsKey(referenceName))
                    _undeclaredNames.Add(referenceName, lineIndex);
            }

            internal void Clear() => _undeclaredNames.Clear();
        }

        private readonly ScriptTranslationIssuesCollector _translationIssues = new ScriptTranslationIssuesCollector();
        public IReadOnlyCollection<string> CollectUndeclaredVariableNames() => _translationIssues._undeclaredNames.Keys;
    }

    [AttributeUsage(AttributeTargets.Assembly)]
    public sealed class EntryRunnerTypeAttribute : Attribute // [assembly: EntryRunnerType(typeof(MyClass))]
    {
        public Func<object> Factory { get; }
        //public Type TargetType { get; }
        //public EntryRunnerTypeAttribute(Type targetType)
        //{
        //    TargetType = targetType ?? throw new ArgumentNullException(nameof(targetType));
        //}
        public EntryRunnerTypeAttribute(Func<object> factory)
        {
            Factory = factory;
        }
    }

    public class ScriptControlConfiguration
    {
        internal readonly string[] _translationSuppression;
        public bool TempEnabled { get; }
        protected string TempDirectoryPath { get; }

        internal readonly string[] NoWarn;

        public ScriptControlConfiguration(bool tempEnabled, string? tempDirectoryPath, bool enabledLoadFromDisk, string[] translationSuppression, string[] noWarn)
        {
            _translationSuppression = translationSuppression ?? throw new ArgumentNullException(nameof(translationSuppression));
            TempEnabled = tempEnabled;
            TempDirectoryPath = tempDirectoryPath ?? "";
            EnabledLoadFromDisk = enabledLoadFromDisk;
            NoWarn = noWarn ?? throw new ArgumentNullException(nameof(noWarn));
        }

        internal bool EnabledLoadFromDisk { get; }

        internal string EnsureFilePathX(string folderName, string fileName)
        {
            string folderPath = Path.Combine(TempDirectoryPath, folderName);
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderPath);
            }
            string filePath = Path.Combine(folderPath, fileName);
            return filePath;
        }

        private string EnsureFilePath(string folderName, string fileName)
        {
            string filePath = EnsureFilePathX(folderName, fileName);
            OnTempFileAdd(filePath);
            return filePath;
        }
        protected virtual void OnTempFileAdd(string filePath)
        {
        }
#pragma warning disable CA1822 // Mark members as static
        protected internal virtual void TempFileWriteAllBytes(string folderName, string fileName, byte[] bytes)
        {
            File.WriteAllBytes(EnsureFilePath(folderName, fileName), bytes);
        }

        protected internal virtual void TempFileWriteAllLines(string folderName, string fileName, IEnumerable<string> contents)
        {
            File.WriteAllLines(EnsureFilePath(folderName, fileName), contents);
        }

        protected internal virtual void TempFileWriteAllText(string folderName, string fileName, string contents)
        {
            File.WriteAllText(EnsureFilePath(folderName, fileName), contents);
        }
#pragma warning restore CA1822 // Mark members as static
    }
}