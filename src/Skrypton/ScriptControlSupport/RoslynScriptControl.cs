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
        internal static UnloadableAssemblyLoadContextContext? CompileCSharpProgram(ScriptControlConfiguration config, string programName, int codeNumber, string csCode, string[] nowarn)
        {
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(csCode);
            MetadataReference[] references = new[]
            {
                MetadataReference.CreateFromFile(Assembly.Load("netstandard").Location), // => C:\Program Files\dotnet\shared\Microsoft.NETCore.App\10.0.0\netstandard.dll *** his is not the compiler’s API reference.
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(typeof(IDisposable).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests).Assembly.Location),
            };

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
                if (!specificDiagnosticOptions.ContainsKey(nowarnItem))
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