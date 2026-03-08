using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.RuntimeSupport;
using Skrypton.Tests.CSharpWriter.CodeTranslation.IntegrationTests;
using Skrypton.RuntimeSupport.Implementations;
using Microsoft.Testing.Platform.Services;

namespace Skrypton.Tests
{
    public abstract class TestBase
    {
        protected const int lineIndex1 = 1;
        protected const int lineIndex2 = 2;
        protected const int lineIndex3 = 3;

        public const string CSFileExtension = ".cs"; // ".cstxt"
        public CultureInfo TestCulture { get; set; } = CultureInfo.InvariantCulture;
        public IRuntimeHost CreateRuntimeHost(IServiceProvider hostServices) => new TestRuntimeHost(hostServices);
        public IRuntimeLogger RuntimeLogger => new TestRuntimeLogger(this);
        public TestContext TestContext { get; set; }
        internal string TestName => this.TestContext!.TestName;
        internal string SaveExpectedActualFiles(string testName, string workItemName
                , string fileName
                , string expected_xml, string actual_xml
            )
        {
            var test_case_name_tokens = this.TestContext.TestName.Split('_');
            string folderpath_tc = workItemName + "/" + test_case_name_tokens.Last();

            SaveContentToFile("expected/" + folderpath_tc, fileName, expected_xml);
            SaveContentToFile("actual/" + folderpath_tc, fileName, actual_xml);

            string expectedDirPath = System.IO.Path.Combine(this.TestContext.TestRunResultsDirectory, "expected");
            string actualDirPath = System.IO.Path.Combine(this.TestContext.TestRunResultsDirectory, "actual");
            string startCommand = "\"C:\\Program Files\\WinMerge\\WinMergeU.exe\" \"" + expectedDirPath + "\" \"" + actualDirPath + "\"";

            return SaveContentToFile(null, "winMergeStarter.bat", startCommand);
        }

        public const char NewLineNormalized = '\n';


        private string SaveContentToFile(string subdir, string fileName, string content)
        {
            //if (this.TestContext != null)
            {
                string subdirPath = this.TestContext.TestRunResultsDirectory;
                if (subdir != null)
                {
                    subdirPath = System.IO.Path.Combine(subdirPath, subdir);
                    System.IO.Directory.CreateDirectory(subdirPath);
                }

                var di = new System.IO.DirectoryInfo(subdirPath);
                if (!di.Exists)
                    di.Create();

                if (fileName.Length > 69)// 69? 27? or 20!
                    throw new InvalidOperationException("File name too long. Length:" + fileName.Length + ", path:" + fileName);

                ///if (fileName.Length > 60)
                ///    fileName = fileName.Substring(0, 60);

                string filePath = System.IO.Path.Combine(subdirPath, fileName);
                if (filePath.Length > 271) // 271? 264? 240!!!
                    throw new InvalidOperationException("File path too long. Length:" + filePath.Length + ", path:" + filePath);

                ///LongFileSupport.WriteAllText(filePath, content);
                System.IO.File.WriteAllText(filePath, content);

                this.TestContext.AddResultFile(filePath);
                return filePath;
            }
        }

        private DefaultRuntimeSupportClassFactory _defaultRuntimeSupportClassFactoryInstance;
        protected DefaultRuntimeSupportClassFactory DefaultRuntimeSupportClassFactoryInstance
        {
            get
            {
                if (_defaultRuntimeSupportClassFactoryInstance == null)
                {
                    _defaultRuntimeSupportClassFactoryInstance = DefaultRuntimeSupportClassFactory.Create(CreateRuntimeHost(CreateTestHostServices()), RuntimeLogger, TestCulture);
                }
                return _defaultRuntimeSupportClassFactoryInstance;
            }
        }

        protected void TestCSharpCodeTranslation(string vbsSource) // TODO remove 'WithoutScaffoldingTranslator'
        {
            string actualCs = DefaultCSharpTranslation.GetTranslatedProgramCode(TestCulture, vbsSource, []);
            string expectCs = TextResourceHelper.LoadResourceText<TestBase>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension);
            TestCSharpCodeTranslationCore(expectCs, actualCs);
        }

        private void TestCSharpCodeTranslationCore(string expectCs, string actualCs) // TODO remove 'WithoutScaffoldingTranslator'
        {
            string chainName = TestName;
            string fileSuffix = CSFileExtension;
            string workItemName = "Script";// TestContext.TestName;
            string text_e = expectCs;
            string text_a = actualCs;
            if (expectCs != null)
            {
                var cmp = string.CompareOrdinal(text_e, text_a);
                if (cmp != 0)
                {
                    string[] arr_expect = text_e.SplitLines().ToArray();
                    string[] arr_actual = text_a.SplitLines().ToArray();

                    int? diffAtIndex = null;
                    if (arr_actual != null)
                    {
                        for (int idx = 0; idx < arr_actual.Length; idx++)
                        {
                            if (idx >= arr_expect.Length)
                            {
                                break;
                            }
                            else
                            {
                                var cmpItem = string.CompareOrdinal(arr_expect[idx], arr_actual[idx]);
                                if (arr_expect[idx].NormalizeLineEndings() != arr_actual[idx].NormalizeLineEndings())
                                {
                                    diffAtIndex = idx;
                                    break;
                                }
                            }
                        }
                    }

                    SaveExpectedActualFiles(chainName, workItemName, chainName + fileSuffix, expectCs, actualCs);

                    string translated_cs_expected = expectCs;
                    string translated_cs_actual = actualCs;
                    int mismatchIndex = FindFirstMismatchIndex(translated_cs_expected, translated_cs_actual, out int mismatchLine, out int mismatchColumn);
                    string snippetE = GetMismatchedSnippet(translated_cs_expected, mismatchIndex, 100);
                    string snippetA = GetMismatchedSnippet(translated_cs_actual, mismatchIndex, 100);
                    string failed_text = $"C# translation failed. Mismatch at line:{mismatchLine}, column:{mismatchColumn} (Index:{mismatchIndex}) \r\nE:'{snippetE}' \r\nA:'{snippetA}'";

                    Assert.Fail(failed_text);// $"File content different at index:{diffAtIndex}");
                }
                else
                {
                }
                return;
            }
            else
            {
                Assert.IsTrue(actualCs == null || actualCs.Length == 0);
            }
        }

        internal static int FindFirstMismatchIndex(string a, string b, out int line, out int column)
        {
            line = 1;
            column = 1;

            int minLength = Math.Min(a.Length, b.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (a[i] != b[i])
                    return i;
                if (a[i] == '\n') // handle windows and unix line endings
                {
                    line++;
                    column = 1;
                }
                else if (a[i] != '\r') // ignore carriage return
                {
                    column++;
                }
            }
            if (a.Length != b.Length)
                return minLength;
            return -1; // no mismatch
        }
        internal static string GetMismatchedSnippet(string s, int startIndex, int maxLength)
        {
            if (startIndex > s.Length)
                return "";
            int endOfLine = s.IndexOfAny(['\r', '\n'], startIndex);
            if (endOfLine == -1)
                endOfLine = s.Length;

            //int remaining  = s.Length - startIndex;
            int take = Math.Min(maxLength, endOfLine - startIndex);
            return s.Substring(startIndex, take);
        }

        internal TestHostServices CreateTestHostServices(Action<TestHostServices> setup = null)
        {
            var container = new TestHostServices();
            setup?.Invoke(container);
            return container;
        }
    }

    internal sealed class TestRuntimeLogger : IRuntimeLogger
    {
        public TestRuntimeLogger(TestBase tst)
        {
        }

        public void LogException(Exception exception)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));
            Console.WriteLine("VBS-Exception:" + exception);
        }
    }

    internal sealed class TestRuntimeHost : IRuntimeHost
    {
        private readonly IServiceProvider _hostServices;

        public TestRuntimeHost(IServiceProvider hostServices)
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        TService IRuntimeHost.TryGetRuntimeHostService<TService>()
        {
            //return _hostServices.GetService(typeof(TService)) as TService;
            return _hostServices.GetService<TService>();
        }
    }

    internal sealed class TestHostServices : IServiceProvider
    {
        private readonly Dictionary<string, Func<object>> _providers = new Dictionary<string, Func<object>>(StringComparer.OrdinalIgnoreCase);
        public TestHostServices()
        {
        }
        internal int ProvidersCount => _providers.Count;
        internal TestHostServices RegisterHostService<T>(Func<T> serviceProvider) where T : class
        {
            _providers.Add(typeof(T).FullName, () => (object)serviceProvider());
            return this;
        }
        public object GetService(Type serviceType)
        {
            if (serviceType == null) throw new ArgumentNullException(nameof(serviceType));
            if (_providers.TryGetValue(serviceType.FullName ?? "", out Func<object> serviceProvider))
            {
                return serviceProvider() ?? throw new InvalidOperationException($"Service '{serviceType.FullName}' factorization failed.");
            }
            throw new NotSupportedException($"Service '{serviceType.FullName}' not registered.");
        }
    }
    internal sealed class TestHostObjectFactoryHostService : IHostObjectFactoryHostService
    {
        public TestHostObjectFactoryHostService()
        {

        }

        private readonly Dictionary<string, Func<IRuntimeHost, object>> _factories = new Dictionary<string, Func<IRuntimeHost, object>>(StringComparer.OrdinalIgnoreCase);
        internal TestHostObjectFactoryHostService RegisterObjectFactory<T>(string progId, Func<IRuntimeHost, T> factory)
        {
            _factories.Add(progId, (h) => factory(h));
            return this;
        }

        public Func<IRuntimeHost, object> TryGetObjectFactoryRegistration(string progId)
        {
            return _factories.TryGetValue(progId, out Func<IRuntimeHost, object> factory) ? factory : null;
        }
    }
}
