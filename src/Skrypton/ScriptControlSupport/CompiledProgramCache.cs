using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// A bounded, least-recently-used cache of compiled script programs (see issue #43). Translating and compiling a script with Roslyn takes a
    /// noticeable amount of time and every compiled program is loaded as a new assembly. Assemblies loaded with Assembly.Load can not be unloaded
    /// (on .NET Framework there is no unloadable AssemblyLoadContext at all), so without this cache a long-running host that executes the same
    /// script over and over again would load one more assembly per call and grow without bound. With the cache, identical programs (the same
    /// translated C# and the same compilation settings) are compiled and loaded once and then reused.
    /// Note: Evicting an entry only means that the program will be compiled again if it is needed again - the evicted assembly stays loaded
    /// (assemblies can not be unloaded), so distinct programs that are not cached still add to the memory used by the process.
    /// </summary>
    public static class CompiledProgramCache
    {
        /// <summary>
        /// The default maximum number of compiled programs that are kept
        /// </summary>
        public const int DefaultMaxSize = 256;

        private static readonly object _lock = new object();
        private static readonly Dictionary<string, LinkedListNode<KeyValuePair<string, Assembly>>> _entries = new Dictionary<string, LinkedListNode<KeyValuePair<string, Assembly>>>(StringComparer.Ordinal);
        private static readonly LinkedList<KeyValuePair<string, Assembly>> _recentlyUsed = new LinkedList<KeyValuePair<string, Assembly>>();
        private static int _maxSize = DefaultMaxSize;

        /// <summary>
        /// The maximum number of compiled programs that are kept (the least recently used program is evicted when this is exceeded). Setting this
        /// to zero disables caching, so that every execution compiles and loads the program again. This must not be negative.
        /// </summary>
        public static int MaxSize
        {
            get
            {
                lock (_lock)
                {
                    return _maxSize;
                }
            }
            set
            {
                if (value < 0)
                    throw new ArgumentOutOfRangeException(nameof(value), "must not be negative");
                lock (_lock)
                {
                    _maxSize = value;
                    TrimToMaxSize();
                }
            }
        }

        /// <summary>
        /// The number of compiled programs that are currently cached
        /// </summary>
        public static int Count
        {
            get
            {
                lock (_lock)
                {
                    return _entries.Count;
                }
            }
        }

        /// <summary>
        /// Remove all cached programs (the assemblies stay loaded, see the class remarks)
        /// </summary>
        public static void Clear()
        {
            lock (_lock)
            {
                _entries.Clear();
                _recentlyUsed.Clear();
            }
        }

        internal static string GetKey(string csCode, IEnumerable<string> noWarn, bool tempEnabled, bool enabledLoadFromDisk, bool treatWarningsAsErrors)
        {
            if (csCode == null)
                throw new ArgumentNullException(nameof(csCode));
            if (noWarn == null)
                throw new ArgumentNullException(nameof(noWarn));

            List<string> sortedNoWarn = new List<string>(noWarn);
            sortedNoWarn.Sort(StringComparer.Ordinal);
            string keySource = string.Join("\n", sortedNoWarn) + "\n" + tempEnabled + "\n" + enabledLoadFromDisk + "\n" + treatWarningsAsErrors + "\n" + csCode;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(keySource));
                StringBuilder hex = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                    hex.Append(b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
                return hex.ToString();
            }
        }

        internal static Assembly GetOrAdd(string key, Func<Assembly> compile)
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));
            if (compile == null)
                throw new ArgumentNullException(nameof(compile));

            lock (_lock)
            {
                if (_entries.TryGetValue(key, out LinkedListNode<KeyValuePair<string, Assembly>>? node))
                {
                    _recentlyUsed.Remove(node);
                    _recentlyUsed.AddFirst(node);
                    return node.Value.Value;
                }
            }

            // Compile outside of the lock (it is slow), if two threads compile the same program at the same time then the first one to finish wins
            Assembly compiled = compile();

            lock (_lock)
            {
                if (_entries.TryGetValue(key, out LinkedListNode<KeyValuePair<string, Assembly>>? existingNode))
                    return existingNode.Value.Value;
                if (_maxSize == 0)
                    return compiled;

                LinkedListNode<KeyValuePair<string, Assembly>> newNode = _recentlyUsed.AddFirst(new KeyValuePair<string, Assembly>(key, compiled));
                _entries.Add(key, newNode);
                TrimToMaxSize();
                return compiled;
            }
        }

        private static void TrimToMaxSize()
        {
            while (_entries.Count > _maxSize)
            {
                LinkedListNode<KeyValuePair<string, Assembly>> leastRecentlyUsed = _recentlyUsed.Last!;
                _recentlyUsed.RemoveLast();
                _entries.Remove(leastRecentlyUsed.Value.Key);
            }
        }
    }
}
