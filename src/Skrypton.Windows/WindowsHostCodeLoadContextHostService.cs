using System;
using System.IO;
using System.Reflection;
#if !NETFRAMEWORK
using System.Runtime.Loader;
#endif
using Skrypton.ScriptControlSupport;

namespace Skrypton.Windows
{
    /// <summary>
    /// The code load context of the compiled programs of IEagerScriptEngine versions (issue #190). On .NET every version is loaded into its own
    /// collectible AssemblyLoadContext, which is unloaded when the version is released (the assembly is collected once nothing references it). On
    /// the .NET Framework there is no collectible load context: the version is loaded into the default context and releasing it only drops the
    /// reference (the assembly stays loaded).
    /// </summary>
    public sealed class WindowsHostCodeLoadContextHostService : IHostCodeLoadContextHostService
    {
        public IHostCodeLoadContext Load(string name, byte[] assemblyImage, byte[] symbols)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Value cannot be null or empty.", nameof(name));
            if (assemblyImage == null) throw new ArgumentNullException(nameof(assemblyImage));
            if (symbols == null) throw new ArgumentNullException(nameof(symbols));
#if NETFRAMEWORK
            return new CodeLoadContext(Assembly.Load(assemblyImage, symbols));
#else
            CollectibleLoadContext loadContext = new CollectibleLoadContext(name);
            using (MemoryStream image = new MemoryStream(assemblyImage, writable: false))
            using (MemoryStream pdb = new MemoryStream(symbols, writable: false))
                return new CodeLoadContext(loadContext.LoadFromStream(image, pdb), loadContext);
#endif
        }

        /// <summary>
        /// The loaded program and (on .NET) its collectible load context. The load context itself must not reference the assembly: the runtime keeps
        /// an unloading context alive until its assemblies are collected, so such a reference would keep both forever.
        /// </summary>
        private sealed class CodeLoadContext : IHostCodeLoadContext
        {
#if NETFRAMEWORK
            public CodeLoadContext(Assembly assembly)
            {
                Assembly = assembly;
            }
#else
            private CollectibleLoadContext? _loadContext;

            public CodeLoadContext(Assembly assembly, CollectibleLoadContext loadContext)
            {
                Assembly = assembly;
                _loadContext = loadContext;
            }
#endif

            public Assembly Assembly { get; }

            public void Unload()
            {
#if NETFRAMEWORK
                // the .NET Framework can not unload an assembly (without an AppDomain): only the reference of the version goes
#else
                System.Threading.Interlocked.Exchange(ref _loadContext, null)?.Unload();
#endif
            }
        }
#if !NETFRAMEWORK

        private sealed class CollectibleLoadContext : AssemblyLoadContext
        {
            // the dependencies of the program (Skrypton, the runtime) come from the load context that loaded Skrypton
            private static readonly AssemblyLoadContext SkryptonLoadContext = GetLoadContext(typeof(IHostCodeLoadContext).Assembly) ?? Default;

            public CollectibleLoadContext(string name)
                : base(name, isCollectible: true)
            {
            }

            protected override Assembly? Load(AssemblyName assemblyName) => SkryptonLoadContext.LoadFromAssemblyName(assemblyName);
        }
#endif
    }
}
