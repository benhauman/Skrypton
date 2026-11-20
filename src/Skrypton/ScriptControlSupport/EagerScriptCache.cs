using System;
using System.Collections.Generic;
using System.Threading;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// The process-wide cache of the compiled versions of IEagerScriptEngine (issue #190). An entry is reference-counted by the compiled code
    /// handles and the engines that use it: a version is never compiled twice while one of them is alive, and when the last one is done the entry is
    /// removed and its program released (unloaded where its code load context can, see IHostCodeLoadContextHostService). Parallel callers of a new
    /// version wait for one compilation; a failed compilation is not cached.
    /// </summary>
    internal static class EagerScriptCache
    {
        internal sealed class Entry
        {
            public Entry(string key, Lazy<CompiledScript> script)
            {
                Key = key;
                Script = script;
            }

            public string Key { get; }
            public Lazy<CompiledScript> Script { get; }
            public int References { get; set; } // guarded by _lock
        }

        private static readonly object _lock = new object();
        private static readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.Ordinal);

        /// <summary>
        /// The compiled version of the key with one more reference (release it with Release) - compiled now unless a handle or an engine still uses
        /// it (reused)
        /// </summary>
        internal static Entry Acquire(string key, Func<CompiledScript> compile, out bool reused)
        {
            Entry? entry;
            lock (_lock)
            {
                reused = _entries.TryGetValue(key, out entry);
                if (!reused)
                {
                    entry = new Entry(key, new Lazy<CompiledScript>(compile, LazyThreadSafetyMode.ExecutionAndPublication));
                    _entries.Add(key, entry);
                }
                entry!.References++;
            }

            try
            {
                _ = entry.Script.Value;
            }
            catch
            {
                // a failed compilation is not cached (Lazy would keep the exception), so the next call compiles again
                lock (_lock)
                {
                    entry.References--;
                    if (_entries.TryGetValue(key, out Entry? failed) && (failed == entry))
                        _entries.Remove(key);
                }
                throw;
            }
            return entry;
        }

        /// <summary>
        /// One more reference to an entry that is in use (the caller holds a reference)
        /// </summary>
        internal static void AddReference(Entry entry)
        {
            lock (_lock)
            {
                if (entry.References <= 0)
                    throw new ObjectDisposedException(nameof(CompiledCodeHandle), "The compiled version was released");
                entry.References++;
            }
        }

        /// <summary>
        /// Release a reference - the last one removes the entry and releases its program
        /// </summary>
        internal static void Release(Entry entry)
        {
            bool last;
            lock (_lock)
            {
                entry.References--;
                last = entry.References == 0;
                if (last && _entries.TryGetValue(entry.Key, out Entry? cached) && (cached == entry))
                    _entries.Remove(entry.Key);
            }
            if (last)
                entry.Script.Value.Unload();
        }
    }
}
