using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.RuntimeSupport;

namespace Skrypton.Tests
{
    public abstract class TestBase
    {
        public CultureInfo TestCulture { get; set; } = CultureInfo.InvariantCulture;
        public TestContext TestContext { get; set; }
        protected string TestName => this.TestContext!.TestName;
        protected void SaveExpectedActualFiles(string testName, string workItemName
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

            SaveContentToFile(null, "winMergeStarter.bat", startCommand);

        }


        private void SaveContentToFile(string subdir, string fileName, string content)
        {
            if (this.TestContext != null)
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
            }
        }

        private DefaultRuntimeSupportClassFactory _defaultRuntimeSupportClassFactoryInstance;
        protected DefaultRuntimeSupportClassFactory DefaultRuntimeSupportClassFactoryInstance
        {
            get
            {
                if (_defaultRuntimeSupportClassFactoryInstance == null)
                {
                    _defaultRuntimeSupportClassFactoryInstance = DefaultRuntimeSupportClassFactory.Create(TestCulture);
                }
                return _defaultRuntimeSupportClassFactoryInstance;
            }
        }

        private static string RenderTranslatedStatement(TranslatedStatement s)
        {
            if (s.IndentationDepth == 0)
                return s.Content;
            string txt = new string(' ', s.IndentationDepth * 4) + s.Content;
            return txt;
        }

        protected void TestCSharpCodeTranslation(string vbsSource)
        {
            string[] output = Skrypton.CSharpWriter.DefaultTranslator
                .Translate(TestCulture, vbsSource, new string[0], renderCommentsAboutUndeclaredVariables: false)
                .Select(s => RenderTranslatedStatement(s))
                .ToArray();

            string expected = TextResourceHelper.LoadResourceText<TestBase>("Skrypton.Tests.VbsResources." + TestName + ".cstxt");

            string chainName = TestName;
            string fileSuffix = ".cstxt";
            //    AreEqualStringArray(TestName, ".cstxt",
            string[] arr_expected = expected.Split(new[] { Environment.NewLine }, StringSplitOptions.None).Select(s => s.Trim()).Where(s => s != "").ToArray();
            string[] arr_actual = output.Select(s => s.Trim()).Where(s => s != "").ToArray();
            string text_a_raw = string.Join("\r\n", output);
            //    );
            //}
            //private void AreEqualStringArray(string chainName, string fileSuffix, string[] arr_expected, string[] arr_actual)
            //{
            string workItemName = "Script";// TestContext.TestName;
            string text_e = arr_expected == null ? null : string.Join("\r\n", arr_expected);
            string text_a = arr_actual == null ? null : string.Join("\r\n", arr_actual);
            if (arr_expected != null)
            {
                if (text_e != text_a)
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
                                if (arr_expected[idx] != arr_actual[idx])
                                {
                                    diffAtIndex = idx;
                                    break;
                                }
                            }
                        }
                    }

                    SaveExpectedActualFiles(chainName, workItemName, chainName + fileSuffix, expected, text_a_raw);
                    Assert.Fail($"File content different at index:{diffAtIndex}");
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
    }
}
