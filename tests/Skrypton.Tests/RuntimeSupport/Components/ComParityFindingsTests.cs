#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.RuntimeSupport.Components;
using Skrypton.ScriptControlSupport;
using Skrypton.Tests.Application;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// One test per finding of the COM parity check against Windows VBScript (see ComParityTests for the black-box comparison)
    /// </summary>
    [TestClass]
    public sealed class ComParityFindingsTests : TestBase
    {
        public sealed class ValueCollector
        {
            public List<object?> Values { get; } = new List<object?>();
            public void Report(object? value) => Values.Add(value);
        }

        /// <summary>A host object whose property returns an object with a default member</summary>
        public sealed class HostWithList
        {
            public sealed class ItemList
            {
                private readonly string[] _items = { "zero", "one" };

                [System.Runtime.InteropServices.DispId(0)]
                [Skrypton.RuntimeSupport.Attributes.IsDefault]
                public string Item(int index) => _items[index];
            }

            public ItemList Items { get; } = new ItemList();
        }

        private List<object?> Execute(string vbs, params (string Name, object Value)[] objects)
        {
            var collector = new ValueCollector();
            IScriptControl scriptengine = CreateScriptControlClass(
                CreateRuntimeHost(CreateTestHostServices(s => s.RegisterHostService<IHostProcessControlHostService>(() => new TestHostProcessControlHostService()))),
                CreateScriptControlConfiguration(false, [], []));
            scriptengine.AddObject("collector", collector, false);
            foreach (var (name, value) in objects)
                scriptengine.AddObject(name, value, false);
            scriptengine.ExecuteStatement(vbs);
            return collector.Values;
        }

        [TestMethod]
        public void MemberCallWithArgumentsUsesTheDefaultMemberOfAParameterLessMember()
        {
            // obj.Items(1) is obj.Items.Item(1), as with VBScript / COM (eg. match.SubMatches(0))
            Assert.AreEqual("one", Execute("collector.Report host.Items(1)\r\n", ("host", new HostWithList()))[0]);
        }

        [TestMethod]
        public void UnmatchedOptionalRegExpGroupIsEmpty()
        {
            var values = Execute(
                "Dim re : Set re = CreateObject(\"VBScript.RegExp\") : re.Pattern = \"(a)(b)?\"\r\n" +
                "Dim m : Set m = re.Execute(\"ac\")(0)\r\n" +
                "collector.Report IsEmpty(m.SubMatches(1))\r\n" +
                "collector.Report \"\" & m\r\n"); // the Match object itself would be passed to Report (as with VBScript); "" & m uses its default Value
            CollectionAssert.AreEqual(new object?[] { true, "a" }, values.ToArray());
        }

        [TestMethod]
        public void ComponentErrorsReachTheScriptWithTheirVBScriptNumber()
        {
            // Errors raised inside a component are not wrapped (TargetInvocationException) - Err.Number is the VBScript number
            var values = Execute(
                "Dim fso : Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n" +
                "On Error Resume Next\r\n" +
                "fso.GetFile \"" + Path.Combine(Path.GetTempPath(), "skrypton-missing-" + Guid.NewGuid().ToString("N")) + "\"\r\n" +
                "collector.Report Err.Number\r\n");
            Assert.AreEqual(53, Convert.ToInt32(values[0]));
        }

        [TestMethod]
        public void AFileThatIsStillOpenCanBeDeleted()
        {
            // VBScript releases an unassigned TextStream immediately; .NET only when it is collected - so the default file system service shares
            // files for deletion
            string file = Path.Combine(Path.GetTempPath(), "skrypton-open-" + Guid.NewGuid().ToString("N") + ".txt");
            var values = Execute(
                "Dim fso : Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n" +
                "fso.CreateTextFile \"" + file + "\"\r\n" +
                "fso.DeleteFile \"" + file + "\"\r\n" +
                "collector.Report fso.FileExists(\"" + file + "\")\r\n");
            Assert.AreEqual(false, values[0]);
        }

        [TestMethod]
        public void ObjectArgumentIsConvertedThroughItsDefaultProperty()
        {
            // A Folder passed where a path is expected is its Path (COM variant conversion)
            var values = Execute(
                "Dim fso : Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n" +
                "collector.Report TypeName(fso.GetFolder(fso.GetSpecialFolder(2)))\r\n");
            Assert.AreEqual("Folder", values[0]);
        }

        [TestMethod]
        [DataRow(@"C:\a\b.txt", "b.txt", "b", "txt", @"C:\a")]
        [DataRow(@"C:\a\dir\", "dir", "dir", "", @"C:\a")]
        [DataRow(@"C:\a", "a", "a", "", @"C:\")]
        [DataRow(@"C:\", "", "", "", "")]
        [DataRow("a/b/c.txt", "c.txt", "c", "txt", "a/b")]
        [DataRow("x.tar.gz", "x.tar.gz", "x.tar", "gz", "")]
        public void FileSystemObjectPathFunctionsTreatBothSeparatorsOnEveryPlatform(string path, string fileName, string baseName, string extension, string parent)
        {
            var fso = (IFileSystem)TestBuiltInComponents.Create("Scripting.FileSystemObject");
            Assert.AreEqual(fileName, fso.GetFileName(path), "GetFileName");
            Assert.AreEqual(baseName, fso.GetBaseName(path), "GetBaseName");
            Assert.AreEqual(extension, fso.GetExtensionName(path), "GetExtensionName");
            Assert.AreEqual(parent, fso.GetParentFolderName(path), "GetParentFolderName");
        }

        [TestMethod]
        [DataRow(@"C:\a", "b.txt", @"C:\a\b.txt")]
        [DataRow(@"C:\a\", "b.txt", @"C:\a\b.txt")]
        [DataRow(@"C:\a", @"\b.txt", @"C:\a\b.txt")]
        [DataRow("a/b", "c.txt", "a/b/c.txt")]
        public void FileSystemObjectBuildPathKeepsTheSeparatorOfThePathOnEveryPlatform(string path, string name, string expected)
        {
            var fso = (IFileSystem)TestBuiltInComponents.Create("Scripting.FileSystemObject");
            Assert.AreEqual(expected, fso.BuildPath(path, name));
        }

        [TestMethod]
        public void ScriptletTypeLibGuidHasTheWindowsFormat()
        {
            var values = Execute("Dim tl : Set tl = CreateObject(\"Scriptlet.TypeLib\")\r\nDim g : g = tl.Guid\r\ncollector.Report g\r\n");
            string guid = (string)values[0]!;
            Assert.HasCount(40, guid.ToCharArray(), "38 characters and two NUL characters");
            StringAssert.Matches(guid.Substring(0, 38), new System.Text.RegularExpressions.Regex(@"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$"));
            Assert.AreEqual("\0\0", guid.Substring(38));
        }

        [TestMethod]
        [DataRow("CreateObject(\"Scripting.FileSystemObject\")", 9, DisplayName = "no default member (the C# ToString fallback does not count)")]
        [DataRow("CreateObject(\"Scripting.Dictionary\")", 9, DisplayName = "default member with a parameter")]
        [DataRow("CreateObject(\"Scriptlet.TypeLib\")", 9, DisplayName = "Scriptlet.TypeLib")]
        [DataRow("CreateObject(\"ADODB.Connection\")", 8, DisplayName = "default ConnectionString")]
        [DataRow("CreateObject(\"Scripting.FileSystemObject\")|GetSpecialFolder(2)", 8, DisplayName = "default Folder.Path")]
        public void VarTypeFollowsAnExplicitParameterLessDefaultMember(string expression, int expected)
        {
            Assert.AreEqual(expected, Convert.ToInt32(Execute(SetObject(expression) + "collector.Report VarType(o)\r\n")[0]));
        }

        /// <summary>
        /// "creation|member" sets o to the creation expression, then (if there is a member) to o.member
        /// </summary>
        private static string SetObject(string expression)
        {
            string[] parts = expression.Split('|');
            return "Dim o : Set o = " + parts[0] + "\r\n" + ((parts.Length > 1) ? "Set o = o." + parts[1] + "\r\n" : "");
        }

        [TestMethod]
        [DataRow("CreateObject(\"VBScript.RegExp\")", "IRegExp2")]
        [DataRow("CreateObject(\"VBScript.RegExp\")|Execute(\"a\")", "IMatchCollection2")]
        [DataRow("CreateObject(\"WScript.Shell\")", "IWshShell3")]
        [DataRow("CreateObject(\"Scriptlet.TypeLib\")", "IGenScriptletTLib")]
        [DataRow("CreateObject(\"Msxml2.ServerXMLHTTP.6.0\")", "IServerXMLHTTPRequest2")]
        [DataRow("CreateObject(\"Shell.Application\")", "IShellDispatch6")]
        [DataRow("GetObject(\"winmgmts:\")", "SWbemServicesEx")]
        [DataRow("CreateObject(\"Scripting.FileSystemObject\")|GetSpecialFolder(2)", "Folder")]
        public void TypeNamesAreTheWindowsInterfaceNames(string expression, string expected)
        {
            Assert.AreEqual(expected, Execute(SetObject(expression) + "collector.Report TypeName(o)\r\n")[0]);
        }
    }
}
