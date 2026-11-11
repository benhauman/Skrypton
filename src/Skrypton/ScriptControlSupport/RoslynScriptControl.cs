using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
//using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Emit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
//using Microsoft.CodeAnalysis.Scripting;

namespace Skrypton.ScriptControlSupport
{

    internal static class RoslynScriptControl
    {
        //private ScriptState<object> _state;
        //private ScriptOptions _options;
        //
        //// Host objects dictionary
        //private Dictionary<string, object> _hostObjects = new Dictionary<string, object>();
        //
        //public RoslynScriptControl()
        //{
        //    //_options = ScriptOptions.Default
        //    //    .WithImports("System")
        //    //    .WithReferences(AppDomain.CurrentDomain.GetAssemblies()
        //    //        .Where(a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location)));
        //    //_state = null;
        //}
        //
        // Add code like AddCode in ScriptControl
        //public async Task AddCodeAsync(string code, CancellationToken cancellationToken)
        //{
        //    if (_state == null)
        //    {
        //        _state = await CSharpScript.RunAsync(code, _options, globals: GetGlobalsObject(), globalsType: null, cancellationToken: cancellationToken).ConfigureAwait(false);
        //    }
        //    else
        //    {
        //        _state = await _state.ContinueWithAsync(code, _options, cancellationToken).ConfigureAwait(false);
        //    }
        //}

        // Execute statement like ExecuteStatement in ScriptControl
        //public async Task ExecuteStatementAsync(string statement, CancellationToken cancellationToken)
        //{
        //    await AddCodeAsync(statement, cancellationToken).ConfigureAwait(false);
        //}
        //
        //// Run a function with parameters like Run in ScriptControl
        //public async Task<object> RunAsync(string functionName, params object[] args)
        //{
        //    // Generate call codeExpression dynamically
        //    var argList = string.Join(",", args.Select((a, i) => $"args[{i}]").ToArray());
        //    var code = $"{functionName}({argList})";
        //    var globals = GetGlobalsObject();
        //    dynamic result = await CSharpScript.EvaluateAsync(code, _options, globals: globals);
        //    return result;
        //}
        //
        //// Add host object like AddObject
        //public void AddObject(string name, object obj, bool addMembers = false)
        //{
        //    if (_hostObjects.ContainsKey(name))
        //        throw new InvalidOperationException($"An object with this name '{name}' already exists in the script namespace.");
        //
        //    if (addMembers)
        //    {
        //        // Flatten members to globals dynamically
        //        foreach (var prop in obj.GetType().GetProperties())
        //        {
        //            _hostObjects[prop.Name] = prop.GetValue(obj);
        //        }
        //        foreach (var method in obj.GetType().GetMethods().Where(m => !m.IsSpecialName))
        //        {
        //            _hostObjects[method.Name] = method.CreateDelegate(typeof(Delegate), obj);
        //        }
        //    }
        //    else
        //    {
        //        _hostObjects[name] = obj;
        //    }
        //}
        //
        //private dynamic GetGlobalsObject()
        //{
        //    dynamic expando = new ExpandoObject();
        //    var dict = (IDictionary<string, object>)expando;
        //    foreach (var kvp in _hostObjects)
        //        dict[kvp.Key] = kvp.Value;
        //    return expando;
        //}
        //
        private static readonly Lazy<MetadataReference[]> _compilationReferences = new Lazy<MetadataReference[]>(BuildCompilationReferences);

        internal static MetadataReference[] GetCompilationReferences() => _compilationReferences.Value;

        /// <summary>
        /// The facade assemblies that the generated program needs: Skrypton is a netstandard2.0 library, so "netstandard" is required, and on
        /// .NET (Core) the core types are exposed through "System.Runtime". These are loaded by their full names because on .NET Framework a load by
        /// simple name is not resolved against the GAC (FileNotFoundException: Could not load file or assembly 'netstandard', see issue #44).
        /// </summary>
        internal static readonly string[] FacadeAssemblyNames =
        {
            "netstandard, Version=2.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51",
            "System.Runtime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
        };

        private static MetadataReference[] BuildCompilationReferences()
        {
            var locations = new List<string>();
            foreach (string facadeAssemblyName in FacadeAssemblyNames)
            {
                string? location = TryGetFacadeAssemblyLocation(new AssemblyName(facadeAssemblyName));
                if (location != null)
                    locations.Add(location);
            }
            locations.Add(typeof(IDisposable).Assembly.Location);
            locations.Add(typeof(object).Assembly.Location);
            locations.Add(typeof(Console).Assembly.Location);
            locations.Add(typeof(Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests).Assembly.Location);

            if (!locations.Any(l => Path.GetFileName(l).Equals("netstandard.dll", StringComparison.OrdinalIgnoreCase)))
            {
                throw new FileNotFoundException(
                    "The 'netstandard' facade assembly that is required to compile translated scripts could not be found (tried '" + FacadeAssemblyNames[0] +
                    "', the simple name 'netstandard' and the 'Facades' folder of the runtime directory '" + System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory() + "')"
                );
            }
            return locations
                .Where(l => !string.IsNullOrEmpty(l))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(l => (MetadataReference)MetadataReference.CreateFromFile(l))
                .ToArray();
        }

        private static string? TryGetFacadeAssemblyLocation(AssemblyName assemblyName)
        {
            // 1. By full name (this is what works on .NET Framework, where netstandard.dll is in the GAC)
            string? location = TryLoadLocation(() => Assembly.Load(assemblyName));
            if (location != null)
                return location;

            // 2. By simple name (on .NET (Core) the runtime resolves this to the version that it ships with)
            location = TryLoadLocation(() => Assembly.Load(assemblyName.Name!));
            if (location != null)
                return location;

            // 3. The "Facades" folder of the .NET Framework runtime directory (eg. C:\Windows\Microsoft.NET\Framework\v4.0.30319\Facades\netstandard.dll)
            string facadePath = Path.Combine(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "Facades", assemblyName.Name + ".dll");
            return File.Exists(facadePath) ? facadePath : null;
        }

        private static string? TryLoadLocation(Func<Assembly> load)
        {
            try
            {
                string location = load().Location;
                return string.IsNullOrEmpty(location) ? null : location;
            }
            catch (FileNotFoundException)
            {
                return null;
            }
            catch (FileLoadException)
            {
                return null;
            }
            catch (BadImageFormatException)
            {
                return null;
            }
        }

        internal static UnloadableAssemblyLoadContextContext? CompileCSharpProgram(ScriptControlConfiguration config, string programName, int codeNumber, string csCode, string[] nowarn)
        {
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(csCode);
            MetadataReference[] references = GetCompilationReferences();

            /*
            string dotnetRoot = Path.GetDirectoryName(typeof(object).Assembly.Location);// path to dotnet runtime: eg 'C:\Program Files\dotnet\shared\Microsoft.NETCore.App\10.0.3'
            //string netstandardRefDir = Path.Combine(
            //    dotnetRoot,
            //    "packs",
            //    "NETStandard.Library.Ref",
            //    "2.0.0",
            //    "ref",
            //    "netstandard2.0"
            //); // eg 'C:\Program Files\dotnet\shared\Microsoft.NETCore.App\10.0.3\packs\NETStandard.Library.Ref\2.0.0\ref\netstandard2.0'
            string netstandardRefDir = dotnetRoot;

            if (!Directory.Exists(netstandardRefDir))
                throw new DirectoryNotFoundException($"netstandard2.0 reference pack cannot be resolved:'{netstandardRefDir}'");


            // STEP 1: Load .NET reference assemblies (required for Assembly.Load)
            var referencesX = Directory.GetFiles(netstandardRefDir, "*.dll")
                .Select(f => MetadataReference.CreateFromFile(f))
                .Cast<MetadataReference>()
                .OrderBy(x => x.Display)
                .ToList();


            references = references.Concat(referencesX).Distinct().ToArray();
            */

            Dictionary<string, ReportDiagnostic> specificDiagnosticOptions = new Dictionary<string, ReportDiagnostic>();
            //specificDiagnosticOptions["CS0219"] = ReportDiagnostic.Suppress; // error CS0219: The variable 'ForWriting' is assigned but its value is never used

            foreach (string nowarnItem in nowarn)
            {
                if (!specificDiagnosticOptions.ContainsKey(nowarnItem))// TryAdd, ReportDiagnostic.Suppress))
                {
                    specificDiagnosticOptions.Add(nowarnItem, ReportDiagnostic.Suppress);
                }
            }
            string codeName = $"InMemDynAsmKey{codeNumber}";
            string tempFolderName = $"programName{codeNumber}";
            string fileNameCsc = $"{codeName}.cs";
            string fileNameDll = $"{codeName}.dll";
            string fileNamePdb = $"{codeName}.pdb";
            string fileNameErr = $"errors.log";

            if (config.TempEnabled)
            {
                config.TempFileWriteAllText(tempFolderName, fileNameCsc, csCode);
            }

            // Compilation options (warnings as errors, warning level 4)
            CSharpCompilationOptions options = new CSharpCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary, // 'DynamicallyLinkedLibrary' or '
                warningLevel: 4,
                generalDiagnosticOption: ReportDiagnostic.Error,
                specificDiagnosticOptions: specificDiagnosticOptions,
                optimizationLevel: OptimizationLevel.Debug, // Keep debug info
                allowUnsafe: false
            );


            CSharpCompilation compilation = CSharpCompilation.Create(
                assemblyName: $"InMemDynAsmKey{codeNumber}",
                syntaxTrees: new[] { syntaxTree },
                references: references,
                options: options
            );

            using MemoryStream peStream = new MemoryStream();
            using MemoryStream pdbStream = new MemoryStream();

            // Emit with debug info
            EmitResult emitResult = compilation.Emit(
                peStream,
                pdbStream,
                options: new EmitOptions(debugInformationFormat: DebugInformationFormat.PortablePdb,
                            pdbFilePath: fileNamePdb
                    )
            );

            if (!string.IsNullOrEmpty(fileNameErr))
            {
                config.TempFileWriteAllLines(tempFolderName, fileNameErr!, emitResult.Diagnostics.Select(d => d.ToString()));
            }
            // Equivalent to results.Errors
            if (!emitResult.Success)
            {
                StringBuilder errorsBuffer = new StringBuilder();

                foreach (Diagnostic diagnostic in emitResult.Diagnostics)
                {
                    if (diagnostic.Severity == DiagnosticSeverity.Error)
                    {
                        errorsBuffer.AppendLine($"c# {diagnostic.Severity} suppressed: {diagnostic.IsSuppressed} IsWarningAsError:{diagnostic.IsWarningAsError} :" + diagnostic.ToString());
                    }
                    else
                    {
                        errorsBuffer.AppendLine($"c# {diagnostic.Severity} suppressed: {diagnostic.IsSuppressed} IsWarningAsError:{diagnostic.IsWarningAsError}::" + diagnostic.ToString());
                    }
                }

                foreach (Diagnostic diagnostic in emitResult.Diagnostics)
                    Console.WriteLine("cs " + diagnostic);

                Console.WriteLine(errorsBuffer.ToString());

                // In unit tests, you can fail like this:
                throw new CompilationFailedException("Compilation failed.");
                // Or if using NUnit/xUnit:
                // Assert.Fail("Compilation failed.");
            }
            peStream.Seek(0, SeekOrigin.Begin);
            pdbStream.Seek(0, SeekOrigin.Begin);

            // Load assembly from memory
            byte[] assemblyBytes = peStream.ToArray();
            byte[] pdbBytes = peStream.ToArray();

            // write the .dll
            if (config.TempEnabled && !string.IsNullOrEmpty(fileNameDll))
            {
                config.TempFileWriteAllBytes(tempFolderName, fileNameDll!, assemblyBytes);
            }
            // write the .pdb
            if (config.TempEnabled && !string.IsNullOrEmpty(fileNamePdb))
            {
                config.TempFileWriteAllBytes(tempFolderName, fileNamePdb, pdbBytes);
            }

            Func<Assembly> asmLoad;
            if (!string.IsNullOrEmpty(fileNameDll) && config.TempEnabled && !config.EnabledLoadFromDisk)
            {
                string dllFilePath = config.EnsureFilePathX(tempFolderName, fileNameDll);
                if (!File.Exists(dllFilePath))
                    throw new InvalidOperationException("generated program could not be found.");
                System.Console.WriteLine("Load assembly from:" + dllFilePath);
                var asm = Assembly.LoadFile(dllFilePath); // this a dynamic assembly => do not use Assembly.Load(AssemblyName.GetAssemblyName(dllFilePath))!!! otherwise 'System.IO.FileNotFoundException'
                asmLoad = () => asm;
            }
            else
            {
                asmLoad = () => Assembly.Load(assemblyBytes);
            }
            // return Assembly.Load(peStream.ToArray());
            UnloadableAssemblyLoadContextContext context = new UnloadableAssemblyLoadContextContext(asmLoad());
            //context.LoadedAssembly = context.LoadFromStream(new MemoryStream(assemblyBytes));
            //context.LoadFromAssemblyPath
            //return context;
            //context.LoadedAssembly = Assembly.Load(assemblyBytes);

            // var type = asm.GetType("MyClass1");
            // var method = type.GetMethod("MyMethod1", BindingFlags.Public | BindingFlags.Static);
            // var del = (Func<int, int>)method.CreateDelegate(typeof(Func<int, int>));
            //
            return context;
        }
    }

    internal sealed class UnloadableAssemblyLoadContextContext// : System.Runtime.Loader.AssemblyLoadContext
    {
        public Assembly LoadedAssembly { get; set; }

        public UnloadableAssemblyLoadContextContext(Assembly loadedAssembly)// : base(isCollectible: true)
        {
            LoadedAssembly = loadedAssembly ?? throw new ArgumentNullException(nameof(loadedAssembly));
        }
        //protected override Assembly Load(AssemblyName assemblyName)
        //{
        //    return base.Load(assemblyName);
        //}
        //
        //protected override nint LoadUnmanagedDll(string unmanagedDllName)
        //{
        //    return base.LoadUnmanagedDll(unmanagedDllName);
        //}

#pragma warning disable CA1822 // Mark members as static
        internal void UnloadContextCollectAndWait()
#pragma warning restore CA1822 // Mark members as static
        {
            //this.Unload();
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }
    }
}