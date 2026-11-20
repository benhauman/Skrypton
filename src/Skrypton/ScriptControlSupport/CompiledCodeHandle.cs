using System;
using System.Collections.Generic;
using System.Globalization;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// A compiled version of an IEagerScriptEngine script (issue #190), returned by CompileCode. It creates engines of the version without passing
    /// the code again and is thread-safe. Dispose it to retire the version: once the last engine created from it is disposed too, the version is
    /// released (on .NET its code load context is unloaded, on the .NET Framework only the reference goes).
    /// </summary>
    public sealed class CompiledCodeHandle : IDisposable
    {
        private readonly object _lock = new object();
        private EagerScriptCache.Entry? _entry; // null once disposed, so that a disposed handle does not keep the program
        private readonly CultureInfo _culture;
        private readonly ScriptEngineConfiguration _configuration;
        private readonly IReadOnlyDictionary<string, Type> _declaredObjects;

        internal CompiledCodeHandle(EagerScriptCache.Entry entry, bool reused, CultureInfo culture, ScriptEngineConfiguration configuration, IReadOnlyDictionary<string, Type> declaredObjects)
        {
            _entry = entry;
            Key = entry.Key;
            Procedures = entry.Script.Value.Procedures;
            Reused = reused;
            _culture = culture;
            _configuration = configuration;
            _declaredObjects = declaredObjects;
        }

        /// <summary>
        /// The cache key of the version: a SHA-256 of the code, the culture, the declared objects (names, types, addMembers) and the options
        /// </summary>
        public string Key { get; }

        /// <summary>
        /// Whether CompileCode reused the version (no parse, no translation, no compile) because a handle or an engine of it was still alive
        /// </summary>
        public bool Reused { get; }

        /// <summary>
        /// The Subs and Functions of the compiled script
        /// </summary>
        public Procedures Procedures { get; }

        /// <summary>
        /// A new engine of this version, ready to bind the declared objects and call procedures (phase 4) - dispose it when the run is done
        /// </summary>
        public IEagerScriptEngine CreateEngine(IRuntimeHost runtimeHost, IRuntimeLogger runtimeLogger)
        {
            if (runtimeHost == null) throw new ArgumentNullException(nameof(runtimeHost));
            if (runtimeLogger == null) throw new ArgumentNullException(nameof(runtimeLogger));
            EagerScriptCache.Entry entry;
            lock (_lock)
            {
                entry = _entry ?? throw new ObjectDisposedException(nameof(CompiledCodeHandle));
                EagerScriptCache.AddReference(entry);
            }
            return new EagerScriptEngine(runtimeHost, runtimeLogger, _culture, _configuration, entry, _declaredObjects);
        }

        public void Dispose()
        {
            EagerScriptCache.Entry? entry;
            lock (_lock)
            {
                entry = _entry;
                _entry = null;
            }
            if (entry != null)
                EagerScriptCache.Release(entry);
        }
    }
}
