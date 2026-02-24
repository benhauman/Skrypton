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

namespace Skrypton.Tests
{
    public abstract class TestBase
    {
        public const string CSFileExtension = ".cs"; // ".cstxt"
        public CultureInfo TestCulture { get; set; } = CultureInfo.InvariantCulture;
        public IRuntimeLogger RuntimeLogger => new TestRuntimeLogger(this);
        public TestContext TestContext { get; set; }
        protected string TestName => this.TestContext!.TestName;
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
                    _defaultRuntimeSupportClassFactoryInstance = DefaultRuntimeSupportClassFactory.Create(RuntimeLogger, TestCulture);
                }
                return _defaultRuntimeSupportClassFactoryInstance;
            }
        }

        protected void TestCSharpCodeTranslation(string vbsSource) // TODO remove 'WithoutScaffoldingTranslator'
        {
            string[] output = DefaultCSharpTranslation.GetTranslatedStatements(TestCulture, vbsSource, []);

            string expectedCs = TextResourceHelper.LoadResourceText<TestBase>("Skrypton.Tests.VbsResources." + TestName + CSFileExtension);
            string[] arr_expected = expectedCs.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray();
            string[] arr_actual = output.Select(s => s.Trim()).Where(s => s != "").ToArray();
            string text_a_raw = string.Join(NewLineNormalized, output);
            TestCSharpCodeTranslationCore(vbsSource, text_a_raw, arr_expected, arr_actual, expectedCs);
        }
        //protected void TestCSharpCodeTranslationWithoutScaffoldingTranslator(string vbsSource, string[] arr_expected) // TODO remove 'WithoutScaffoldingTranslator'
        //{
        //    string expectedCs = string.Join(NewLineNormalized, arr_expected);
        //    //myAssert.AreEqual(
        //    //expected.Select(s => s.Trim()).ToArray(),
        //    var output = WithoutScaffoldingTranslator.GetTranslatedStatements(TestCulture, vbsSource, WithoutScaffoldingTranslator.DefaultConsoleExternalDependencies);
        //    //);
        //    string[] arr_actual = output.Select(s => s.Trim()).Where(s => s != "").ToArray();
        //    string text_a_raw = string.Join(NewLineNormalized, output);
        //    TestCSharpCodeTranslationCore(vbsSource, text_a_raw, arr_expected, arr_actual, expectedCs);
        //}
        private void TestCSharpCodeTranslationCore(string vbsSource, string text_a_raw, string[] arr_expected, string[] arr_actual, string expectedCs) // TODO remove 'WithoutScaffoldingTranslator'
        {

            string chainName = TestName;
            string fileSuffix = CSFileExtension;
            //    AreEqualStringArray(TestName, CSFileExtension,
            //string[] arr_expected = expectedCs.SplitLines().Select(s => s.Trim()).Where(s => s != "").ToArray();
            //string[] arr_actual = output.Select(s => s.Trim()).Where(s => s != "").ToArray();
            //string text_a_raw = string.Join(NewLineNormalized, output);
            //    );
            //}
            //private void AreEqualStringArray(string chainName, string fileSuffix, string[] arr_expected, string[] arr_actual)
            //{
            string workItemName = "Script";// TestContext.TestName;
            string text_e = arr_expected == null ? null : string.Join(NewLineNormalized, arr_expected).NormalizeLineEndings();
            string text_a = arr_actual == null ? null : string.Join(NewLineNormalized, arr_actual).NormalizeLineEndings();
            if (arr_expected != null)
            {
                var cmp = string.CompareOrdinal(text_e, text_a);
                if (cmp != 0)
                {
                    int? diffAtIndex = null;
                    if (arr_actual != null)
                    {
                        for (int idx = 0; idx < arr_actual.Length; idx++)
                        {
                            if (idx >= arr_expected.Length)
                            {
                                break;
                            }
                            else
                            {
                                var cmpItem = string.CompareOrdinal(arr_expected[idx], arr_actual[idx]);
                                if (arr_expected[idx].NormalizeLineEndings() != arr_actual[idx].NormalizeLineEndings())
                                {
                                    diffAtIndex = idx;
                                    break;
                                }
                            }
                        }
                    }

                    SaveExpectedActualFiles(chainName, workItemName, chainName + fileSuffix, expectedCs, text_a_raw);

                    string translated_cs_expected = expectedCs;
                    string translated_cs_actual = text_a_raw;
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
                Assert.IsTrue(arr_actual == null || arr_actual.Length == 0);
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
            int endOfLine = s.IndexOfAny(new char[] { '\r', '\n' }, startIndex);
            if (endOfLine == -1)
                endOfLine = s.Length;

            //int remaining  = s.Length - startIndex;
            int take = Math.Min(maxLength, endOfLine - startIndex);
            return s.Substring(startIndex, take);
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
}
