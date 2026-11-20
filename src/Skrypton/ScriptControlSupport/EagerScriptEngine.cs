using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// IEagerScriptEngine (issue #190): declare the host objects, set the code, compile it once (process-wide, see EagerScriptCache), then bind the
    /// host objects and call procedures - the global code runs once per engine and the module state is kept between calls. Requires a
    /// ScriptEngineConfiguration in ScriptEngineMode.EagerScript. Not thread-safe: one engine per request (CompiledCodeHandle.CreateEngine).
    /// </summary>
    public sealed class EagerScriptEngine : IEagerScriptEngine
    {
        private enum Phase
        {
            Declare,
            Code,
            Sealed,
            Running,
            Faulted,
            Disposed,
        }

        private readonly IRuntimeHost _runtimeHost;
        private readonly IRuntimeLogger _runtimeLogger;
        private readonly CultureInfo _culture;
        private readonly ScriptEngineConfiguration _configuration;
        private readonly ScriptControlError _error = new ScriptControlError();
        private readonly List<(string Name, Type Type, bool AddMembers)> _declarations = new List<(string Name, Type Type, bool AddMembers)>();
        private readonly Dictionary<string, object> _boundObjects = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private IReadOnlyDictionary<string, Type> _declaredObjects = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase);
        private Phase _phase;
        private string? _code;
        private EagerScriptCache.Entry? _entry;
        private CompiledScriptRun? _run;

        public EagerScriptEngine(IRuntimeHost runtimeHost, IRuntimeLogger runtimeLogger, CultureInfo culture, ScriptEngineConfiguration configuration)
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
            _runtimeLogger = runtimeLogger ?? throw new ArgumentNullException(nameof(runtimeLogger));
            _culture = culture ?? throw new ArgumentNullException(nameof(culture));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            if (configuration.Mode != ScriptEngineMode.EagerScript)
                throw new ArgumentException($"An {nameof(EagerScriptEngine)} needs a configuration in {nameof(ScriptEngineMode)}.{nameof(ScriptEngineMode.EagerScript)}", nameof(configuration));
            _phase = Phase.Declare;
        }

        /// <summary>
        /// An engine of a compiled version (CompiledCodeHandle.CreateEngine) - it starts in phase 4 and holds a reference of the version
        /// </summary>
        internal EagerScriptEngine(IRuntimeHost runtimeHost, IRuntimeLogger runtimeLogger, CultureInfo culture, ScriptEngineConfiguration configuration, EagerScriptCache.Entry entry, IReadOnlyDictionary<string, Type> declaredObjects)
        {
            _runtimeHost = runtimeHost;
            _runtimeLogger = runtimeLogger;
            _culture = culture;
            _configuration = configuration;
            _entry = entry;
            _declaredObjects = declaredObjects;
            _phase = Phase.Sealed;
        }

        public void DeclareObject(string name, Type type, bool addMembers = false)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Value cannot be null or empty.", nameof(name));
            if (type == null) throw new ArgumentNullException(nameof(type));
            Expect(Phase.Declare, nameof(DeclareObject));
            if (_declarations.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"An object with this name is already declared: '{name}'");
            _declarations.Add((name, type, addMembers));
        }

        public void SetCode(string code)
        {
            if (string.IsNullOrEmpty(code)) throw new ArgumentException("Value cannot be null or empty.", nameof(code));
            Expect(Phase.Declare, nameof(SetCode));
            _code = code;
            _phase = Phase.Code;
        }

        public CompiledCodeHandle CompileCode()
        {
            Expect(Phase.Code, nameof(CompileCode));
            _error.Clear();
            Dictionary<string, Type> declaredObjects = _declarations.ToDictionary(d => d.Name, d => d.Type, StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string[]> members = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach ((string name, Type type, bool addMembers) in _declarations)
            {
                // as IScriptControl.AddObject with AddMembers, but from the declared type (there is no instance)
                string[] methods = addMembers
                    ? type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).Where(m => m.DeclaringType != typeof(object)).Select(m => m.Name).Distinct().ToArray()
                    : Array.Empty<string>();
                if (methods.Length > 0)
                    members.Add(name, methods);
            }
            CompiledScriptSource source = new CompiledScriptSource(_code!, _culture, _configuration, declaredObjects.Keys.ToArray(), members);
            string key = GetKey(source);
            IHostCodeLoadContextHostService? codeLoadContextService = _runtimeHost.TryGetRuntimeHostService<IHostCodeLoadContextHostService>();

            EagerScriptCache.Entry entry;
            bool reused;
            try
            {
                entry = EagerScriptCache.Acquire(key, () => CompiledScript.Compile(key, source, codeLoadContextService), out reused);
            }
            catch (Exception e)
            {
                _error.SetFrom(e, null, null, _code);
                throw;
            }
            // one reference for the engine (released by Dispose) and one for the handle
            EagerScriptCache.AddReference(entry);
            _entry = entry;
            _declaredObjects = declaredObjects;
            _phase = Phase.Sealed;
            return new CompiledCodeHandle(entry, reused, _culture, _configuration, declaredObjects);
        }

        public void BindObject(string name, object instance)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Value cannot be null or empty.", nameof(name));
            if (instance == null) throw new ArgumentNullException(nameof(instance));
            Expect(Phase.Sealed, nameof(BindObject));
            if (!_declaredObjects.TryGetValue(name, out Type? declaredType))
                throw new ArgumentException($"The object '{name}' was not declared", nameof(name));
            if (!declaredType.IsInstanceOfType(instance)) // an interface passes for an implementing class or a COM object that implements it
                throw new ArgumentException($"The object '{name}' must be a {declaredType.FullName}, not a {instance.GetType().FullName}", nameof(instance));
            if (_boundObjects.ContainsKey(name))
                throw new InvalidOperationException($"The object '{name}' is already bound");
            _boundObjects.Add(name, instance);
        }

        public Procedures Procedures
        {
            get
            {
                ExpectCompiled(nameof(Procedures));
                return _entry!.Script.Value.Procedures;
            }
        }

        public bool ContainsProcedure(string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            ExpectCompiled(nameof(ContainsProcedure));
            return _entry!.Script.Value.Procedures.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        public void RunGlobalCode()
        {
            Expect(Phase.Sealed, nameof(RunGlobalCode));
            _error.Clear();
            StartRun();
        }

        public object? CallProcedure(string name, params object?[] args)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Value cannot be null or empty.", nameof(name));
            object?[] arguments = args ?? Array.Empty<object?>();
            if ((_phase != Phase.Sealed) && (_phase != Phase.Running))
                throw OutOfOrder(nameof(CallProcedure));
            if (!ContainsProcedure(name))
                throw new ArgumentException($"The script has no Sub or Function '{name}'", nameof(name));
            _error.Clear();
            if (_phase == Phase.Sealed)
                StartRun(); // the global code runs once, before the first procedure
            CompiledScript script = _entry!.Script.Value;
            try
            {
                return _run!.Call(name, arguments);
            }
            catch (Exception e)
            {
                _error.SetFrom(e, script.ProgramAssembly, script.SourceLineMap, script.ScriptContent, script.CsCode);
                throw;
            }
        }

        public Error Error => _error;

        public void Dispose()
        {
            if (_phase == Phase.Disposed)
                return;
            _phase = Phase.Disposed;
            _run?.Dispose();
            _run = null;
            _boundObjects.Clear();
            EagerScriptCache.Entry? entry = _entry;
            _entry = null;
            if (entry != null)
                EagerScriptCache.Release(entry);
        }

        private void StartRun()
        {
            string[] unbound = _declaredObjects.Keys.Where(name => !_boundObjects.ContainsKey(name)).ToArray();
            if (unbound.Length > 0)
                throw new InvalidOperationException($"Every declared object must be bound before the script runs, not bound: {string.Join(", ", unbound)}");
            CompiledScript script = _entry!.Script.Value;
            try
            {
                _run = script.Run(CreateCompatLayer(), _boundObjects);
                _phase = Phase.Running;
            }
            catch (Exception e)
            {
                _phase = Phase.Faulted;
                _error.SetFrom(e, script.ProgramAssembly, script.SourceLineMap, script.ScriptContent, script.CsCode);
                throw;
            }
        }

        private DefaultRuntimeFunctionalityProvider CreateCompatLayer()
        {
            DefaultRuntimeSupportClassFactory factory = DefaultRuntimeSupportClassFactory.Create(_runtimeHost, _runtimeLogger, _culture);
            return new DefaultRuntimeFunctionalityProvider(_runtimeHost, _runtimeLogger, factory.DefaultVBScriptValueRetriever, _culture);
        }

        /// <summary>
        /// The key of CompiledScriptCache (code, culture, names, members, options) and the declared types (assembly-qualified name and module MVID,
        /// so that a changed contract type gives a new version) and addMembers
        /// </summary>
        private string GetKey(CompiledScriptSource source)
        {
            StringBuilder key = new StringBuilder("EagerScript.v1\n").Append(CompiledScriptCache.GetKey(source)).Append('\n');
            foreach ((string name, Type type, bool addMembers) in _declarations.OrderBy(d => d.Name.ToUpperInvariant(), StringComparer.Ordinal))
                key.Append(name.ToUpperInvariant()).Append('=').Append(type.AssemblyQualifiedName).Append('|').Append(type.Module.ModuleVersionId.ToString("N")).Append('|').Append(addMembers).Append('\n');
            return CompiledScriptCache.Hash(key.ToString());
        }

        private void Expect(Phase phase, string member)
        {
            if (_phase != phase)
                throw OutOfOrder(member);
        }

        private void ExpectCompiled(string member)
        {
            if ((_phase != Phase.Sealed) && (_phase != Phase.Running) && (_phase != Phase.Faulted))
                throw OutOfOrder(member);
        }

        private InvalidOperationException OutOfOrder(string member)
        {
            if (_phase == Phase.Disposed)
                throw new ObjectDisposedException(nameof(EagerScriptEngine));
            string state = _phase switch
            {
                Phase.Declare => "objects are declared (call SetCode next)",
                Phase.Code => "the code is set (call CompileCode next)",
                Phase.Sealed => "the code is compiled and sealed (bind the objects, then call RunGlobalCode or CallProcedure)",
                Phase.Running => "the global code has run (call procedures)",
                _ => "the global code of the script failed (create a new engine)",
            };
            return new InvalidOperationException($"{member} is not allowed now: {state}");
        }
    }
}
