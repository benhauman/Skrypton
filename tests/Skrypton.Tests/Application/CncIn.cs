using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Helpline.Application.ScriptingModel;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Tests.Application
{
    [TestClass]
    public class CncIn : TestBase
    {
        /*
            SELECT scriptid   = scr.id
                , scripttext = scr.script
             FROM [dbo].[hlsysscript] AS scr
            WHERE scr.active      = 1
              AND scr.objectdefid = 900 -- 900:connectivity
              AND scr.[type]      = 16 -- 16:ScriptTypeConnectivityIn
              AND scr.scriptmode  = 0 -- 0:eScriptMode.ScriptModeWorking
              AND LEN(ISNULL(scr.script,N'')) > 0 -- TODO: CK_
              ;
         */
        /* --> row number is the customer index : 98:PsoShow, 35:DFSnDLNeu
        SELECT [dbname]
               ,[sizebytes_before]
               ,[sizebytes_after]
               ,[hasfilestreamfilegroup]
               ,[sizegb_before]
               ,[sizegb_after]
               ,[deltamb]
           FROM [CustomerAnalytics].[dbo].[_DatabaseStats]

         */
        [TestMethod]
        public void DC_DATA__hlsysscript_cncIN()
        {
            DoCncInTest();
        }
        [TestMethod]
        public void LUNA12_quxDATA__hlsysscript_cncIN()
        {
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.Connectivity);
            DoCncInTest();
        }
        [TestMethod]
        public void CT98__hlsysscript_cncIN()
        {
            ChainsTest.TestScriptChain(this, TestName, ScriptUsageKind.Connectivity);
            DoCncInTest();
        }

        private void DoCncInTest()
        {
            bool mergeSU_called = false;
            Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext = Helpline.Application.ScriptingModel.ApplicationTestContext.Create(ctx =>
            {
                ctx.HandlerMergeSUs = (obj) =>
                {
                    mergeSU_called = true;
                };
            });
            CncJob session = CreateSampleConnectivityJob(cncTestContext);
            CncObj DoExtendWorkflowCaseIdentity = null;
            session.DoExtendWorkflowCaseOverride = (oi) =>
            {
                DoExtendWorkflowCaseIdentity = (CncObj)oi;
            };
            ExecuteTranslatedProgram(RuntimeLogger, TestCulture, TestContext.TestName, new Dictionary<string, object> { { "session", session } });

            // assert
            Assert.IsFalse(mergeSU_called, "mergeSU_called");
            Assert.IsNotNull(DoExtendWorkflowCaseIdentity, nameof(DoExtendWorkflowCaseIdentity));

        }
        internal static void ExecuteTranslatedProgram(IRuntimeLogger runtimeLogger, CultureInfo culture, string chainName, Dictionary<string, object> externalReferences)
        {
            //
            UnloadableAssemblyLoadContextContext asmctx = CompileCSharpProgram(chainName);
            WeakReference weakRef = new WeakReference(asmctx);//, trackResurrection: true);
            {
                Assembly asm = asmctx.LoadedAssembly;

                DefaultRuntimeSupportClassFactory defaultRuntimeSupportClassFactoryInstance = Skrypton.RuntimeSupport.DefaultRuntimeSupportClassFactory.Create(runtimeLogger, culture);
                Skrypton.RuntimeSupport.IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer = CreateDefaultRuntimeFunctionalityProvider(runtimeLogger, defaultRuntimeSupportClassFactoryInstance.DefaultVBScriptValueRetriever, culture);

                Type tRunner = asm.GetType("TranslatedProgram.Runner", true); // TODO: use an assembly attribute for this class instead of reflection
                RunnerBase runner = RunnerBase.CreateRunnerInstanceForType(tRunner, compatLayer);

                EnvironmentReferencesBase environmentReferences = runner.CreateEnvironmentReferencesInstance();

                var properties = environmentReferences.GetType().GetProperties();
                var propertiesNameInfo = properties.ToDictionary(x => x.Name, x => x, StringComparer.OrdinalIgnoreCase);

                foreach (KeyValuePair<string, object> externalReferencesEntry in externalReferences)
                {
                    string externalReferenceName = externalReferencesEntry.Key;
                    object externalReferenceInstance = externalReferencesEntry.Value;
                    environmentReferences.InitializeExternalReference(externalReferenceName, externalReferenceInstance);

                    if (!propertiesNameInfo.TryGetValue(externalReferenceName, out PropertyInfo pi_externalReference1))
                        throw new InvalidOperationException($"Invalid external reference '{externalReferenceName}'.");
                    // sanity check
                    _ = pi_externalReference1.GetValue(environmentReferences);
                }

                runner.Run(environmentReferences);
            }
            asmctx.UnloadContextCollectAndWait();

            if (!weakRef.IsAlive)
                Console.WriteLine("ALC successfully unloaded");
            else
                Console.WriteLine("ALC still alive");
        }
        internal static DefaultRuntimeFunctionalityProvider CreateDefaultRuntimeFunctionalityProvider(IRuntimeLogger runtimeLogger, IAccessValuesUsingVBScriptRules valueRetriever, CultureInfo culture)
        {
            DefaultRuntimeFunctionalityProvider provider = new DefaultRuntimeFunctionalityProvider(runtimeLogger, valueRetriever, culture);
            provider.RegisterObjectCreateFactory("Scripting.Dictionary", () => new Skrypton.Tests.RuntimeSupport.Implementations.MyScriptingDictionaryCpuAny());
            provider.RegisterObjectCreateFactory("Shell.Application", () => new Skrypton.Tests.RuntimeSupport.Implementations.MyShellApplication());
            provider.RegisterObjectCreateFactory("Msxml2.ServerXMLHTTP.6.0", () => new Skrypton.Tests.RuntimeSupport.Implementations.MyServerXMLHTTP60());
            provider.RegisterObjectCreateFactory("Msxml2.DOMDocument", () => new Skrypton.Tests.RuntimeSupport.Implementations.MyMsxml2DOMDocument());
            return provider;
        }

        //class MyDefaultRuntimeFunctionalityProvider : Skrypton.RuntimeSupport.Implementations.DefaultRuntimeFunctionalityProvider
        //{
        //    public MyDefaultRuntimeFunctionalityProvider(Func<string, string> nameRewriter, Skrypton.RuntimeSupport.IAccessValuesUsingVBScriptRules valueRetriever, CultureInfo culture)
        //        : base(valueRetriever, culture)
        //    {
        //    }
        //
        //    //public override object CREATEOBJECT(object value)
        //    //{
        //    //    string progid = (string)value;
        //    //    if (string.Equals(progid, "Scripting.Dictionary", StringComparison.OrdinalIgnoreCase))
        //    //    {
        //    //        return new ScriptingDictionary();
        //    //    }
        //    //    return base.CREATEOBJECT(value);
        //    //}
        //}

        internal static UnloadableAssemblyLoadContextContext CompileCSharpProgram(string chainName)
        {
            string translated_cs_expected = TextResourceHelper.LoadResourceText<CncIn>("Skrypton.Tests.VbsResources." + chainName + CSFileExtension);
            return CompileCSharpProgram(chainName, translated_cs_expected);
        }
        internal static UnloadableAssemblyLoadContextContext CompileCSharpProgram(string chainName, string translated_cs)
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
                generalDiagnosticOption: ReportDiagnostic.Error,
                specificDiagnosticOptions: new Dictionary<string, ReportDiagnostic>
                {
                    ["CS0219"] = ReportDiagnostic.Suppress
                }
            );

            CSharpCompilation compilation = CSharpCompilation.Create(
                "InMemDynAsmKey2",
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
                throw new Exception("Compilation failed.");
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
            context.LoadedAssembly = context.LoadFromStream(new MemoryStream(assemblyBytes));
            //context.LoadFromAssemblyPath
            return context;
        }
        internal sealed class UnloadableAssemblyLoadContextContext : System.Runtime.Loader.AssemblyLoadContext
        {
            public Assembly LoadedAssembly { get; set; }

            public UnloadableAssemblyLoadContextContext() : base(isCollectible: true)
            {
            }
            protected override Assembly Load(AssemblyName assemblyName)
            {
                return base.Load(assemblyName);
            }

            protected override nint LoadUnmanagedDll(string unmanagedDllName)
            {
                return base.LoadUnmanagedDll(unmanagedDllName);
            }

            internal void UnloadContextCollectAndWait()
            {
                this.Unload();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }



        internal static Helpline.Application.ScriptingModel.CncJob CreateSampleConnectivityJob(Helpline.Application.ScriptingModel.IApplicationTestContext cncTestContext)
        {
            return new CncJob(cncTestContext)
            {
                m_cfg = new CncConfigGroup("Root").AddGroup("casetypEs", caseTypes =>
                {
                    caseTypes.AddGroup("type1", t1 =>
                    {
                        t1.InitValue("CaseType", v => { v.m_data = null; });
                        t1.InitValue("MailAttributeKey", v => { v.m_data = "PersonCommunication.PersonEmail_CA.EmailAddress"; });
                        t1.InitValue("Type", v => { v.m_data = "1"; });
                    });
                    caseTypes.AddGroup("type2", t2 =>
                    {
                        t2.InitValue("CaseTyp", v => { v.m_data = null; });
                        t2.InitValue("MailAttributeKey", v => { v.m_data = "PersonCommunication.PersonEmail_CA.EmailAddress"; });
                        t2.InitValue("Type", v => { v.m_data = "-2"; });
                    });
                })
                                    ,

                m_mailRequest = new CncMail()
                {
                    Subject = "this a feedbacl [#20190711-0012]. Awesome",
                    data_From = "peter.pan@wonderland.com"
                }
            };
        }
    }
}
