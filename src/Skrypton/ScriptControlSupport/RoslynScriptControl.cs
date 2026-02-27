using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
//using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Emit;
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
        //    // Generate call expression dynamically
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

        internal static UnloadableAssemblyLoadContextContext CompileCSharpProgram(string translated_cs)
        {
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(translated_cs);
            PortableExecutableReference[] references = new[]
            {
                MetadataReference.CreateFromFile(Assembly.Load("netstandard").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(typeof(IDisposable).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests).Assembly.Location),
            };
            // Compilation options (warnings as errors, warning level 4)
            CSharpCompilationOptions options = new CSharpCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary,
                warningLevel: 4,
                generalDiagnosticOption: ReportDiagnostic.Error
            );

            CSharpCompilation compilation = CSharpCompilation.Create(
                "InMemDynAsmKey1",
                new[] { syntaxTree },
                references,
                options
            );

            using MemoryStream peStream = new MemoryStream();
            using MemoryStream pdbStream = new MemoryStream();

            // Emit with debug info
            EmitResult emitResult = compilation.Emit(
                peStream,
                pdbStream,
                options: new EmitOptions(debugInformationFormat: DebugInformationFormat.PortablePdb)
            );

            // Equivalent to results.Errors
            if (!emitResult.Success)
            {
                StringBuilder errorsBuffer = new StringBuilder();

                foreach (Diagnostic diagnostic in emitResult.Diagnostics)
                {
                    if (diagnostic.Severity == DiagnosticSeverity.Error)
                    {
                        errorsBuffer.AppendLine(diagnostic.ToString());
                    }
                }

                Console.WriteLine(errorsBuffer.ToString());

                // In unit tests, you can fail like this:
                throw new InvalidOperationException("Compilation failed.");
                // Or if using NUnit/xUnit:
                // Assert.Fail("Compilation failed.");
            }


            if (!emitResult.Success)
            {
                foreach (Diagnostic diagnostic in emitResult.Diagnostics)
                    Console.WriteLine(diagnostic);
                return null;
            }

            // Load assembly from memory
            peStream.Seek(0, SeekOrigin.Begin);
            // return Assembly.Load(peStream.ToArray());
            var assemblyBytes = peStream.ToArray();
            var context = new UnloadableAssemblyLoadContextContext();
            //context.LoadedAssembly = context.LoadFromStream(new MemoryStream(assemblyBytes));
            //context.LoadFromAssemblyPath
            //return context;
            context.LoadedAssembly = Assembly.Load(assemblyBytes);

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

        public UnloadableAssemblyLoadContextContext()// : base(isCollectible: true)
        {
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

        internal void UnloadContextCollectAndWait()
        {
            //this.Unload();
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }
    }
}