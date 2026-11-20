using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// What a CompiledScript is compiled from (see issue #187): the whole script text, the culture, the names (and members) of the host objects
    /// and the translator / compiler options of the configuration
    /// </summary>
    internal sealed class CompiledScriptSource
    {
        public CompiledScriptSource(string scriptContent, CultureInfo culture, ScriptControlConfiguration configuration, IReadOnlyCollection<string> externalObjectNames, IReadOnlyDictionary<string, string[]> externalObjectMembers)
        {
            ScriptContent = scriptContent ?? throw new ArgumentNullException(nameof(scriptContent));
            Culture = culture ?? throw new ArgumentNullException(nameof(culture));
            Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            ExternalObjectNames = externalObjectNames ?? throw new ArgumentNullException(nameof(externalObjectNames));
            ExternalObjectMembers = externalObjectMembers ?? throw new ArgumentNullException(nameof(externalObjectMembers));
        }

        public string ScriptContent { get; }
        public CultureInfo Culture { get; }
        public ScriptControlConfiguration Configuration { get; }
        public IReadOnlyCollection<string> ExternalObjectNames { get; }
        public IReadOnlyDictionary<string, string[]> ExternalObjectMembers { get; }
    }

    /// <summary>
    /// A process-wide cache of whole compiled scripts (see issue #187), keyed by a hash of the script SOURCE and the options (not of the translated
    /// C#), so an unchanged script is parsed, translated and compiled once and then reused for procedure reflection and execution. Parallel callers
    /// of the same new version wait for one compilation, a failed compilation is not cached. The cache is bounded by CompiledProgramCache.MaxSize
    /// (least recently used entries are evicted, so the previous version of a changed script goes once it is no longer used); runs that still use
    /// an evicted script keep it. As for CompiledProgramCache, compiled assemblies can not be unloaded.
    /// </summary>
    internal static class CompiledScriptCache
    {
        private sealed class Entry
        {
            public Entry(string key, Lazy<CompiledScript> script)
            {
                Key = key;
                Script = script;
            }

            public string Key { get; }
            public Lazy<CompiledScript> Script { get; }
        }

        private static readonly object _lock = new object();
        private static readonly Dictionary<string, LinkedListNode<Entry>> _entries = new Dictionary<string, LinkedListNode<Entry>>(StringComparer.Ordinal);
        private static readonly LinkedList<Entry> _recentlyUsed = new LinkedList<Entry>();
        private static readonly Dictionary<string, ScriptControlProcedures> _proceduresByContent = new Dictionary<string, ScriptControlProcedures>(StringComparer.Ordinal);

        /// <summary>
        /// Return the compiled script of this source and whether this call compiled it (false: it was reused, without parsing, translating or
        /// compiling) - from the cache if the same source (text, culture, host objects and options) was compiled before, otherwise compile it now
        /// </summary>
        internal static CompiledScript GetOrCompile(CompiledScriptSource source, out bool compiled)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            string key = GetKey(source);
            int maxSize = CompiledProgramCache.MaxSize;
            Lazy<CompiledScript> script;
            compiled = false;
            lock (_lock)
            {
                if (_entries.TryGetValue(key, out LinkedListNode<Entry>? node))
                {
                    _recentlyUsed.Remove(node);
                    _recentlyUsed.AddFirst(node);
                    script = node.Value.Script;
                }
                else
                {
                    compiled = true;
                    script = new Lazy<CompiledScript>(() => CompiledScript.Compile(key, source), LazyThreadSafetyMode.ExecutionAndPublication);
                    _entries.Add(key, _recentlyUsed.AddFirst(new Entry(key, script)));
                }
                TrimToMaxSize(maxSize);
            }

            CompiledScript compiledScript;
            try
            {
                compiledScript = script.Value;
            }
            catch
            {
                // a failed compilation is not cached (Lazy would keep the exception), so the next call compiles again
                lock (_lock)
                {
                    if (_entries.TryGetValue(key, out LinkedListNode<Entry>? failedNode) && (failedNode.Value.Script == script))
                        Remove(failedNode);
                }
                throw;
            }
            if (compiled)
            {
                lock (_lock)
                {
                    AddProcedures(GetContentKey(source.Culture, source.ScriptContent), (ScriptControlProcedures)compiledScript.Procedures, maxSize);
                }
            }
            return compiledScript;
        }

        /// <summary>
        /// The procedures of a script text, parsed once per text and culture (IScriptControl.Procedures, issue #187)
        /// </summary>
        internal static ScriptControlProcedures GetProcedures(CultureInfo culture, string scriptContent)
        {
            string contentKey = GetContentKey(culture, scriptContent);
            lock (_lock)
            {
                if (_proceduresByContent.TryGetValue(contentKey, out ScriptControlProcedures? cached))
                    return cached;
            }
            ScriptControlProcedures procedures = ScriptControlProcedures.FromScript(culture, scriptContent);
            lock (_lock)
            {
                AddProcedures(contentKey, procedures, CompiledProgramCache.MaxSize);
            }
            return procedures;
        }

        private static void AddProcedures(string contentKey, ScriptControlProcedures procedures, int maxSize)
        {
            if (maxSize == 0)
            {
                _proceduresByContent.Clear();
                return;
            }
            if (_proceduresByContent.Count >= maxSize)
                _proceduresByContent.Clear(); // bounded; a removed text is parsed again when needed
            _proceduresByContent[contentKey] = procedures;
        }

        internal static string GetKey(CompiledScriptSource source)
        {
            ScriptControlConfiguration config = source.Configuration;
            StringBuilder keySource = new StringBuilder();
            void Add(string value) => keySource.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value).Append('\n');
            Add("CompiledScript.v1");
            Add(source.Culture.Name);
            Add(string.Join(",", source.ExternalObjectNames.Select(n => n.ToUpperInvariant()).OrderBy(n => n, StringComparer.Ordinal)));
            Add(string.Join(";", source.ExternalObjectMembers
                .OrderBy(m => m.Key.ToUpperInvariant(), StringComparer.Ordinal)
                .Select(m => m.Key.ToUpperInvariant() + "=" + string.Join(",", m.Value.OrderBy(v => v, StringComparer.Ordinal)))));
            Add(string.Join(",", config._translationSuppression.OrderBy(v => v, StringComparer.Ordinal)));
            Add(string.Join(",", config.NoWarn.OrderBy(v => v, StringComparer.Ordinal)));
            Add(config.TempEnabled + "," + config.EnabledLoadFromDisk + "," + config.TreatWarningsAsErrors + "," + config.UndeclaredVariablesAreErrors);
            Add(source.ScriptContent);
            return Hash(keySource.ToString());
        }

        private static string GetContentKey(CultureInfo culture, string scriptContent) => Hash(culture.Name + "\n" + scriptContent);

        internal static string Hash(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
                StringBuilder hex = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                    hex.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return hex.ToString();
            }
        }

        private static void Remove(LinkedListNode<Entry> node)
        {
            _recentlyUsed.Remove(node);
            _entries.Remove(node.Value.Key);
        }

        private static void TrimToMaxSize(int maxSize)
        {
            // the entry of the current call stays until its compilation finished (its Lazy is held by the caller anyway)
            while (_entries.Count > maxSize)
                Remove(_recentlyUsed.Last!);
            if (maxSize == 0)
                _proceduresByContent.Clear();
        }
    }
}
