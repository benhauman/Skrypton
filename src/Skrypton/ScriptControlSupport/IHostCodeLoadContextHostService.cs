using System.Reflection;

namespace Skrypton.ScriptControlSupport
{
    /// <summary>
    /// The host service that loads the compiled program of an IEagerScriptEngine version into a code load context (issue #190). Skrypton.Windows
    /// registers one: on .NET a collectible AssemblyLoadContext, so a retired version is unloaded; on the .NET Framework the default context, where
    /// unloading only drops the reference (the assembly stays loaded). Without a registration the program is loaded into the default context.
    /// </summary>
    public interface IHostCodeLoadContextHostService
    {
        /// <summary>
        /// Load the compiled program (its image and its portable PDB, so that runtime errors keep their line numbers)
        /// </summary>
        IHostCodeLoadContext Load(string name, byte[] assemblyImage, byte[] symbols);
    }

    /// <summary>
    /// A compiled program loaded by an IHostCodeLoadContextHostService
    /// </summary>
    public interface IHostCodeLoadContext
    {
        Assembly Assembly { get; }

        /// <summary>
        /// Called once, when no engine and no handle uses the version any more - the assembly is collected once nothing references it (where the
        /// platform can unload assemblies)
        /// </summary>
        void Unload();
    }
}
