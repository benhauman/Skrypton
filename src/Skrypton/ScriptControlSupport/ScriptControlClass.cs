using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.CSharpWriter.Lists;
using Skrypton.LegacyParser.CodeBlocks;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.ScriptControlSupport
{
    public sealed class ScriptControlClass : IScriptControl
    {
        private int _timeout = -1; // -1 means no timeout (infinite execution time)
        private bool _allowUI = false;// Disable UI by default for security reasons (MessageBox, InputBox, etc.)
        private string _language = "";// ('VBScript') No scripting language is selected by default. It must be set explicitly, otherwise script execution will fail
        private ScriptControlStates _state = ScriptControlStates.Initialized;
        private int _sitehWnd = 0; // required (Must be a valid HWND (window handle)) if allowUI is true otherwise ignored. No host window is associated by default => 0
        private bool _useSafeSubset = true; // 'true': Script runs in safe modeUse safe subset of the scripting language (if supported). 'false': Use full language features. Safe subset is used by default. Potentially dangerous objects and operations are blocked
        private Error _error = null; // Default value: null (or Nothing in VBScript) — no error has occurred yet.
#pragma warning disable CS0414 // The field is assigned but its value is never used
        private object _codeObject = null; // Default value: null (or Nothing in VBScript) — no code object has been set yet. This allows you to interact with script members directly, instead of using Run or ExecuteStatement.
#pragma warning restore CS0414 // The field is assigned but its value is never used

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

        Error IScriptControl.Error => _error;

        object IScriptControl.CodeObject => throw new NotImplementedException();

        Procedures IScriptControl.Procedures => throw new NotImplementedException();

        void IScriptControl._AboutBox()
        {
            throw new NotImplementedException();
        }

        private readonly Dictionary<string, object> _addedObjects = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        void IScriptControl.AddObject(string objectName, object objectInstance, bool addMembers)
        {
            if (objectName == null) throw new ArgumentNullException(nameof(objectName));
            if (string.IsNullOrEmpty(objectName)) throw new ArgumentException("Value cannot be null or empty.", nameof(objectName));

            if (_addedObjects.ContainsKey(objectName))
            {
                // “An object with this name already exists in the script namespace.”
                throw new InvalidOperationException($"An object with this name already exists in the script namespace. '{objectName}' HR:0x800A03EC (SCRIPT_E_DUPLICATEOBJECTNAME)"); // SCRIPT_E_DUPLICATEOBJECTNAME
            }
            _addedObjects[objectName] = objectInstance ?? throw new ArgumentNullException(nameof(objectInstance));

            if (addMembers)
            {
                throw new NotImplementedException();
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
            _code.AppendLine(code);
        }

        object IScriptControl.Eval(string Expression)
        {
            throw new NotImplementedException();
        }

        void IScriptControl.ExecuteStatement(string statement)
        {
            if (string.IsNullOrEmpty(statement))
                throw new ArgumentException("Value cannot be null or empty.", nameof(statement));
            //RoslynScriptControl sc = StartAsync(statement, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            //sc.ExecuteStatementAsync(statement)
            string csCode = GenerateCSharpCode(statement);
            //RoslynScriptControl sc = new RoslynScriptControl();
            //sc.ExecuteStatementAsync(csCode, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            UnloadableAssemblyLoadContextContext asmctx = RoslynScriptControl.CompileCSharpProgram(csCode);
            try
            {
                DefaultRuntimeSupportClassFactory defaultRuntimeSupportClassFactoryInstance = Skrypton.RuntimeSupport.DefaultRuntimeSupportClassFactory.Create(EngineRuntimeLogger, EngineCulture);
                using Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer = CreateDefaultRuntimeFunctionalityProvider(defaultRuntimeSupportClassFactoryInstance.RuntimeLogger, defaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever, EngineCulture);
                Type tRunner = asmctx.LoadedAssembly.GetType("TranslatedProgram.Runner", true);
                RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(tRunner, compatLayer);

                EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();

                foreach (KeyValuePair<string, object> externalReferencesEntry in _addedObjects)
                {
                    environmentReferences.InitializeExternalReference(externalReferencesEntry.Key, externalReferencesEntry.Value);
                }

                var globalRefs = runner.Run(environmentReferences);
            }
            finally
            {
                asmctx.UnloadContextCollectAndWait();
                asmctx = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        internal static DefaultRuntimeFunctionalityProvider CreateDefaultRuntimeFunctionalityProvider(IRuntimeLogger runtimeLogger, IAccessValuesUsingVBScriptRules valueRetriever, CultureInfo culture)
        {
            DefaultRuntimeFunctionalityProvider provider = new DefaultRuntimeFunctionalityProvider(runtimeLogger, valueRetriever, culture);
            //provider.RegisterObjectCreateFactory();
            //provider.RegisterObjectCreateFactory();
            return provider;
        }

        object IScriptControl.Run(string procedureName, ref object[] parameters)
        {
            //RoslynScriptControl sc = StartAsync(null, cancellationToken: default).ConfigureAwait(false).GetAwaiter().GetResult();
            throw new NotImplementedException();
        }

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

        private string GenerateCSharpCode(string statementOrNull)
        {
            string scriptContent = _code.ToString(); // Assume this is populated with the script code to be parsed
            if (!string.IsNullOrEmpty(scriptContent))
            {
                scriptContent += "\r\n";
            }
            scriptContent += statementOrNull;
            IReadOnlyCollection<ICodeBlock> parsedBlocks = Skrypton.LegacyParser.Parser.Parse(EngineCulture, scriptContent);


            //var csLines = DefaultCSharpTranslation.GetTranslatedStatements(tst.TestCulture, scriptContent, externalDependencies);
            NonNullImmutableList<string> externalDependencies = _addedObjects.Keys.ToArray().ToNonNullImmutableList();
            var warningLogger = Skrypton.CSharpWriter.DefaultTranslator.CommentsLogger(true, new Skrypton.CSharpWriter.DefaultTranslator.DelegateWrappingWarningLogger(warningMessageText =>
            {
                Console.WriteLine(warningMessageText);
            }));
            NonNullImmutableList<TranslatedStatement> translatedStatements = Skrypton.CSharpWriter.DefaultTranslator.TranslateCore(EngineCulture, scriptContent, externalDependencies, CSharpWriter.CodeTranslation.BlockTranslators.OuterScopeBlockTranslator.OutputTypeOptions.Executable, warningLogger);

            string[] translatedStatementLines = translatedStatements.Select(ts => ts.Content).ToArray();

            string csCode = "";// "[assembly: Skrypton.ScriptControlSupport.EntryRunnerType(() => new TranslatedProgram.Runner())]\r\n";
            csCode += string.Join("\r\n", translatedStatementLines);
            return csCode;
        }
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
}