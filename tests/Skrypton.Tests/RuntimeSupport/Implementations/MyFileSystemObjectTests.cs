using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Skrypton.Tests.RuntimeSupport.Implementations.FileSystemSupport;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    [TestClass]
    public sealed class MyFileSystemObjectTests : TestBase
    {
        private string _root;

        [TestInitialize]
        public void Init()
        {
            _root = Path.Combine(Path.GetTempPath(), "FSO_AllTests_" + Guid.NewGuid());
            Directory.CreateDirectory(_root);
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_root))
                {
                    Directory.Delete(_root, recursive: true);
                }
            }
            catch
            {
                // Swallow cleanup errors to avoid test flakiness on CI (locked handles, etc.).
            }
        }

        // ---------- MyFileSystemObject tests ----------

        [TestMethod]
        public void CreateFolder_ShouldCreateDirectory()
        {
            var fso = new MyFileSystemObject();
            string folderPath = Path.Combine(_root, "SubFolder");

            var folder = fso.CreateFolder(folderPath);

            Assert.IsTrue(Directory.Exists(folderPath));
            Assert.AreEqual("SubFolder", folder.Name);
        }

        [TestMethod]
        public void FileExists_ShouldReturnTrue_WhenFileExists()
        {
            var fso = new MyFileSystemObject();
            string filePath = Path.Combine(_root, "test.txt");
            File.WriteAllText(filePath, "Hello");

            Assert.IsTrue(fso.FileExists(filePath));
            Assert.IsFalse(fso.FileExists(Path.Combine(_root, "missing.txt")));
        }

        [TestMethod]
        public void GetFile_ShouldReturnCorrectFile()
        {
            var fso = new MyFileSystemObject();
            string filePath = Path.Combine(_root, "data.txt");
            File.WriteAllText(filePath, "ABC");

            IFile file = fso.GetFile(filePath);

            Assert.AreEqual("data.txt", file.Name);
            Assert.AreEqual(3L, file.Size);
        }

        [TestMethod]
        public void CreateTextFile_ShouldCreateAndWrite()
        {
            var fso = new MyFileSystemObject();
            string filePath = Path.Combine(_root, "newfile.txt");

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
            var fso = new MyFileSystemObject();
            string filePath = Path.Combine(_root, "open.txt");

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
            File.WriteAllText(Path.Combine(_root, "a.txt"), "A");
            File.WriteAllText(Path.Combine(_root, "b.txt"), "BB");

            var folder = new MyFolder(_root);
            var files = folder.Files;

            Assert.HasCount(2, files);
            Assert.IsTrue(files.Any(f => f.Name == "a.txt"));
            Assert.IsTrue(files.Any(f => f.Name == "b.txt"));
        }

        [TestMethod]
        public void Folder_ShouldReturnCorrectSubFolders()
        {
            Directory.CreateDirectory(Path.Combine(_root, "A"));
            Directory.CreateDirectory(Path.Combine(_root, "B"));

            var folder = new MyFolder(_root);
            var subs = folder.SubFolders;

            Assert.HasCount(2, subs);
            Assert.IsTrue(subs.Any(f => f.Name == "A"));
            Assert.IsTrue(subs.Any(f => f.Name == "B"));
        }

        [TestMethod]
        public void Folder_Copy_ShouldCopyContent()
        {
            // Arrange
            Directory.CreateDirectory(Path.Combine(_root, "src"));
            File.WriteAllText(Path.Combine(_root, "src", "source.txt"), "X");

            var fso = new MyFileSystemObject();
            var src = fso.GetFolder(Path.Combine(_root, "src"));
            //var src = new MyFolder();
            string dst = Path.Combine(_root, "dst");

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
            string from = Path.Combine(_root, "moveFrom");
            string to = Path.Combine(_root, "moveTo");
            Directory.CreateDirectory(from);
            File.WriteAllText(Path.Combine(from, "f.txt"), "x");

            var folder = new MyFolder(from);

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
            string dir = Path.Combine(_root, "del");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "x.txt"), "x");

            var folder = new MyFolder(dir);
            folder.Delete(force: true);

            Assert.IsFalse(Directory.Exists(dir));
        }

        // ---------- MyFile tests ----------

        [TestMethod]
        public void File_ShouldReturnCorrectProperties()
        {
            string filePath = Path.Combine(_root, "info.txt");
            File.WriteAllText(filePath, "1234");

            var f = new MyFile(filePath);

            Assert.AreEqual("info.txt", f.Name);
            Assert.AreEqual(4L, f.Size);
            Assert.AreEqual("File", f.Type);
        }

        [TestMethod]
        public void File_Move_ShouldMoveFile()
        {
            string filePath = Path.Combine(_root, "move.txt");
            File.WriteAllText(filePath, "Data");

            var f = new MyFile(filePath);

            string newPath = Path.Combine(_root, "moved.txt");
            f.Move(newPath);

            Assert.IsFalse(File.Exists(filePath));
            Assert.IsTrue(File.Exists(newPath));
            Assert.AreEqual("Data", File.ReadAllText(newPath));
        }

        [TestMethod]
        public void File_Copy_ShouldCopyFile()
        {
            string filePath = Path.Combine(_root, "copy.txt");
            File.WriteAllText(filePath, "OK");

            var f = new MyFile(filePath);

            string dest = Path.Combine(_root, "copy2.txt");
            f.Copy(dest, overwrite: true);

            Assert.IsTrue(File.Exists(dest));
            Assert.AreEqual("OK", File.ReadAllText(dest));
        }

        [TestMethod]
        public void File_Delete_ShouldRemoveFile()
        {
            string filePath = Path.Combine(_root, "del.txt");
            File.WriteAllText(filePath, "del");

            var f = new MyFile(filePath);
            f.Delete(force: true);

            Assert.IsFalse(File.Exists(filePath));
        }

        // ---------- MyTextStream tests ----------

        [TestMethod]
        public void TextStream_Write_And_Read()
        {
            string filePath = Path.Combine(_root, "ts.txt");

            {
                var ts = new MyTextStream(filePath, IOMode.ForWriting);
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
                var ts = new MyTextStream(filePath, IOMode.ForReading);
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
            string file = Path.Combine(_root, "lines.txt");
            File.WriteAllLines(file, new[] { "A", "B", "C" });

            {
                var ts = new MyTextStream(file, IOMode.ForReading);
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
            string file = Path.Combine(_root, "append.txt");
            File.WriteAllText(file, "X");

            {
                var ts = new MyTextStream(file, IOMode.ForAppending);
                try
                {
                    ts.Write("Y");
                    ts.WriteLine("Z");
                }
                finally
                {
                    ts.Close();
                }
            }

            var content = File.ReadAllText(file).Replace("\r\n", "\n");
            Assert.AreEqual("XYZ\n", content);
        }
    }
}
