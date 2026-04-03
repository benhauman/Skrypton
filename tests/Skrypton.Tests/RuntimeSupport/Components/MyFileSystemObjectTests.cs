using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.Tests.RuntimeSupport.Components.FileSystemSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass]
    public sealed class MyFileSystemObjectTests : TestBase
    {
        private sealed class RootPathInfo
        {
            public string RootPath { get; }
            public string RootName { get; }

            public RootPathInfo(string rootPath, string rootName)
            {
                RootPath = rootPath;
                RootName = rootName;
            }
        }
        private RootPathInfo EnsureRootPath()
        {
            // Path.GetTempPath()
            const string DT_FMT_ID = "yyyyMMdd_hhmmss_fff";
            string rootDir = Path.Combine(this.TestRunResultsDirectory, "FS_" + DateTime.UtcNow.ToString(DT_FMT_ID), TestName);
            Console.WriteLine($"rootDir:{rootDir}");
            var di = Directory.CreateDirectory(rootDir);
            return new RootPathInfo(rootDir, di.Name);
        }

        // ---------- MyFileSystemObject tests ----------

        [TestMethod]
        public void CreateFolder_ShouldCreateDirectory()
        {
            var root = EnsureRootPath();
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            string folderPath = Path.Combine(root.RootPath, "SubFolder");

            var folder = fso.CreateFolder(folderPath);

            Assert.IsTrue(Directory.Exists(folderPath));
            Assert.AreEqual("SubFolder", folder.Name);
        }

        [TestMethod]
        public void FileExists_ShouldReturnTrue_WhenFileExists()
        {
            var root = EnsureRootPath();
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            string filePath = Path.Combine(root.RootPath, "test.txt");
            File.WriteAllText(filePath, "Hello");

            Assert.IsTrue(fso.FileExists(filePath));
            Assert.IsFalse(fso.FileExists(Path.Combine(root.RootPath, "missing.txt")));
        }

        [TestMethod]
        public void GetFile_ShouldReturnCorrectFile()
        {
            var root = EnsureRootPath();
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            string filePath = Path.Combine(root.RootPath, "data.txt");
            File.WriteAllText(filePath, "ABC");

            IFile file = fso.GetFile(filePath);

            Assert.AreEqual("data.txt", file.Name);
            //Assert.AreEqual(3L, file.Size);
        }

        [TestMethod]
        public void CreateTextFile_ShouldCreateAndWrite()
        {
            var root = EnsureRootPath();
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            string filePath = Path.Combine(root.RootPath, "newfile.txt");

            {
                var stream = fso.CreateTextFile(filePath);
                try
                {
                    stream.WriteLine("Hello!");
                }
                finally
                {
                    stream.Close();
                }
            }

            Assert.IsTrue(File.Exists(filePath));
            var normalized = File.ReadAllText(filePath).Replace("\r\n", "\n");
            Assert.AreEqual("Hello!\n", normalized);
        }

        [TestMethod]
        public void OpenTextFile_ForReadingAfterWriting_ShouldReadBack()
        {
            var root = EnsureRootPath();
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            string filePath = Path.Combine(root.RootPath, "open.txt");

            {
                var ts = fso.CreateTextFile(filePath);
                try
                {
                    ts.Write("First");
                    ts.WriteLine();
                    ts.Write("Second");
                }
                finally
                {
                    ts.Close();
                }
            }

            {
                var ts = fso.OpenTextFile(filePath, IOMode.ForReading);
                try
                {
                    var line1 = ts.ReadLine();
                    var rest = ts.ReadAll();
                    Assert.AreEqual("First", line1);
                    Assert.AreEqual("Second", rest);
                }
                finally
                {
                    ts.Close();
                }
            }
        }

        // ---------- MyFolder tests ----------

        [TestMethod]
        public void Folder_ShouldReturnCorrectFiles()
        {
            var root = EnsureRootPath();
            File.WriteAllText(Path.Combine(root.RootPath, "a.txt"), "A");
            File.WriteAllText(Path.Combine(root.RootPath, "b.txt"), "BB");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var folder = fso.GetFolder(root.RootPath);
            var files = folder.Files;

            Assert.HasCount(2, files);
            Assert.IsTrue(files.Any(f => f.Name == "a.txt"));
            Assert.IsTrue(files.Any(f => f.Name == "b.txt"));
        }

        [TestMethod]
        public void Folder_ShouldReturnCorrectSubFolders()
        {
            var root = EnsureRootPath();
            Directory.CreateDirectory(Path.Combine(root.RootPath, "A"));
            Directory.CreateDirectory(Path.Combine(root.RootPath, "B"));

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var folder = fso.GetFolder(root.RootPath);
            var subs = folder.SubFolders;

            Assert.HasCount(2, subs);
            Assert.IsTrue(subs.Any(f => f.Name == "A"));
            Assert.IsTrue(subs.Any(f => f.Name == "B"));
        }

        [TestMethod]
        public void Folder_Copy_ShouldCopyContent()
        {
            // Arrange
            var root = EnsureRootPath();
            Directory.CreateDirectory(Path.Combine(root.RootPath, "src"));
            File.WriteAllText(Path.Combine(root.RootPath, "src", "source.txt"), "X");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var src = fso.GetFolder(Path.Combine(root.RootPath, "src"));
            //var src = new MyFolder();
            string dst = Path.Combine(root.RootPath, "dst");

            // Act
            src.Copy(dst, overwrite: true);

            // Assert
            Assert.IsTrue(File.Exists(Path.Combine(dst, "source.txt")));
            Assert.AreEqual("X", File.ReadAllText(Path.Combine(dst, "source.txt")));
        }

        [TestMethod]
        public void Folder_Move_ShouldMoveDirectory()
        {
            // Arrange
            var root = EnsureRootPath();
            string from = Path.Combine(root.RootPath, "moveFrom");
            string to = Path.Combine(root.RootPath, "moveTo");
            Directory.CreateDirectory(from);
            File.WriteAllText(Path.Combine(from, "f.txt"), "x");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var folder = fso.GetFolder(from);

            // Act
            folder.Move(to);

            // Assert
            Assert.IsFalse(Directory.Exists(from));
            Assert.IsTrue(Directory.Exists(to));
            Assert.IsTrue(File.Exists(Path.Combine(to, "f.txt")));
        }

        [TestMethod]
        public void Folder_Delete_ShouldRemoveDirectory()
        {
            var root = EnsureRootPath();
            string dir = Path.Combine(root.RootPath, "del");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "x.txt"), "x");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var folder = fso.GetFolder(dir);
            folder.Delete(force: true);

            Assert.IsFalse(Directory.Exists(dir));
        }

        // ---------- MyFile tests ----------

        [TestMethod]
        public void File_ShouldReturnCorrectProperties()
        {
            var root = EnsureRootPath();
            string filePath = Path.Combine(root.RootPath, "info.txt");
            File.WriteAllText(filePath, "1234");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var f = fso.GetFile(filePath);

            Assert.AreEqual("info.txt", f.Name);
            //Assert.AreEqual(4L, f.Size);
            //Assert.AreEqual("File", f.Type);
        }

        [TestMethod]
        public void File_Move_ShouldMoveFile()
        {
            var root = EnsureRootPath();
            string filePath = Path.Combine(root.RootPath, "move.txt");
            File.WriteAllText(filePath, "Data");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var f = fso.GetFile(filePath);

            string newPath = Path.Combine(root.RootPath, "moved.txt");
            f.Move(newPath);

            Assert.IsFalse(File.Exists(filePath));
            Assert.IsTrue(File.Exists(newPath));
            Assert.AreEqual("Data", File.ReadAllText(newPath));
        }

        [TestMethod]
        public void File_Copy_ShouldCopyFile()
        {
            var root = EnsureRootPath();
            string filePath = Path.Combine(root.RootPath, "copy.txt");
            File.WriteAllText(filePath, "OK");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var f = fso.GetFile(filePath);

            string dest = Path.Combine(root.RootPath, "copy2.txt");
            f.Copy(dest, overwrite: true);

            Assert.IsTrue(File.Exists(dest));
            Assert.AreEqual("OK", File.ReadAllText(dest));
        }

        [TestMethod]
        public void File_Delete_ShouldRemoveFile()
        {
            var root = EnsureRootPath();
            string filePath = Path.Combine(root.RootPath, "del.txt");
            File.WriteAllText(filePath, "del");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            var f = fso.GetFile(filePath);
            f.Delete(force: true);

            Assert.IsFalse(File.Exists(filePath));
        }

        // ---------- MyTextStream tests ----------

        [TestMethod]
        public void TextStream_Write_And_Read()
        {
            var root = EnsureRootPath();
            string filePath = Path.Combine(root.RootPath, "ts.txt");
            var fso = new MyFileSystemObject(CreateWindowsFileSystem());

            {
                var ts = fso.OpenTextFile(filePath, IOMode.ForWriting, create: true);
                try
                {
                    ts.WriteLine("Line1");
                    ts.Write("Line2");
                }
                finally
                {
                    ts.Close();
                }
            }

            {
                var ts = fso.OpenTextFile(filePath, IOMode.ForReading);
                try
                {
                    string line1 = ts.ReadLine();
                    string rest = ts.ReadAll();

                    Assert.AreEqual("Line1", line1);
                    Assert.AreEqual("Line2", rest);
                }
                finally
                {
                    ts.Close();
                }
            }
        }

        [TestMethod]
        public void TextStream_ReadLine_AtEndOfStream_ShouldBeTrue()
        {
            var root = EnsureRootPath();
            string file = Path.Combine(root.RootPath, "lines.txt");
            File.WriteAllLines(file, new[] { "A", "B", "C" });

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            {
                var ts = fso.OpenTextFile(file, IOMode.ForReading);
                try
                {
                    Assert.AreEqual("A", ts.ReadLine());
                    Assert.AreEqual("B", ts.ReadLine());
                    Assert.AreEqual("C", ts.ReadLine());
                    Assert.IsTrue(ts.AtEndOfStream);
                }
                finally
                {
                    ts.Close();
                }
            }
        }

        [TestMethod]
        public void TextStream_AppendMode_ShouldAppend()
        {
            var root = EnsureRootPath();
            string file = Path.Combine(root.RootPath, "append.txt");
            File.WriteAllText(file, "X");

            var fso = new MyFileSystemObject(CreateWindowsFileSystem());
            {
                var ts = fso.OpenTextFile(file, IOMode.ForAppending);
                try
                {
                    ts.Write("Y");
                    ts.Write("Z");
                }
                finally
                {
                    ts.Close();
                }
            }

            var content = File.ReadAllText(file);
            Assert.AreEqual("XYZ", content);
        }
    }
}
