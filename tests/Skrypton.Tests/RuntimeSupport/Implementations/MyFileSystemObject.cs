using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations.FileSystemSupport
{
    [SourceClassName("Dictionary")] // for TYPENAME(CreateObject("Scripting.Dictionary"))
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[DefaultMember("Item")]
    internal sealed class MyFileSystemObject : IReflectOnClrType, IFileSystem
    {
        /* COM:"Microsoft Scripting Runtime" => 'Interop.Scripting.dll' => C:\Windows\System32\scrrun.dll */
        public bool FileExists(string path)
        {
            return File.Exists(path);
        }
        public bool FolderExists(string path)
        {
            return Directory.Exists(path);
        }
        public bool DriveExists(string path)
        {
            return Directory.Exists(path);
        }

        public object GetDrive(string path)
        {
            string root = Path.GetPathRoot(path);
            if (string.IsNullOrEmpty(root))
                return false;

            // Normalize like "C:\"
            root = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

            var nfo = DriveInfo.GetDrives()
                .Any(d => string.Equals(
                    d.Name,
                    root,
                    StringComparison.OrdinalIgnoreCase));

            throw new NotImplementedException(path);
        }
        public IFile GetFile(string path)
        {
            //throw new NotImplementedException(path);
            if (!File.Exists(path))
                throw new FileNotFoundException(path);
            return new MyFile(path);
        }

        public IFolder GetFolder(string path)
        {
            //throw new NotImplementedException(path);
            if (!Directory.Exists(path))
                throw new DirectoryNotFoundException(path);
            return new MyFolder(path);
        }

        public IFolder CreateFolder(string path)
        {
            _ = nameof(MyFolder);
            //throw new NotImplementedException(path);
            Directory.CreateDirectory(path);
            return new MyFolder(path);
        }

        public ITextStream OpenTextFile(string path, IOMode mode, bool create = false, Tristate format = Tristate.UseDefault)
        //public MyTextStream OpenTextFile(string path, IOMode mode)
        {
            return new MyTextStream(path, mode);
        }

        public ITextStream CreateTextFile(string path, bool overwrite = false, bool unicode = false)
        //public MyTextStream CreateTextFile(string path, bool overwrite = false)
        {
            //throw new NotImplementedException(path);
            if (!overwrite && File.Exists(path))
                throw new IOException("File exists: " + path);
            return new MyTextStream(path, IOMode.ForWriting);
        }
    }
    internal sealed class MyFolder : IFolder
    {
        public string Path { get; }
        public string Name => new DirectoryInfo(Path).Name;

        public MyFolder(string path) => Path = path;

        public long Size
        {
            get
            {
                long sum = 0;
                foreach (var f in Directory.GetFiles(Path, "*", SearchOption.AllDirectories))
                    sum += new FileInfo(f).Length;
                return sum;
            }
        }

        public IReadOnlyCollection<IFile> Files =>
            Directory.GetFiles(Path)
                .Select(f => new MyFile(f))
                .ToArray();

        public IReadOnlyCollection<IFolder> SubFolders =>
            Directory.GetDirectories(Path)
                .Select(d => new MyFolder(d))
                .ToArray();

        public void Delete(bool force)
        {
            Directory.Delete(Path, recursive: force);
        }

        public void Move(string newPath)
        {
            Directory.Move(Path, newPath);
        }

        public void Copy(string newPath, bool overwrite)
        {
            Directory.CreateDirectory(newPath);
            foreach (var file in Directory.GetFiles(Path))
            {
                var dest = System.IO.Path.Combine(newPath, System.IO.Path.GetFileName(file));
                File.Copy(file, dest, overwrite);
            }
        }
    }

    internal sealed class MyFile : IFile
    {
        public string Path { get; }
        public string Name => System.IO.Path.GetFileName(Path);
        public long Size => new FileInfo(Path).Length;
        public string Type => "File";

        public MyFile(string path) => Path = path;

        public void Delete(bool force)
        {
            File.Delete(Path);
        }

        public void Move(string newPath)
        {
            File.Move(Path, newPath);
        }

        public void Copy(string newPath, bool overwrite)
        {
            File.Copy(Path, newPath, overwrite);
        }

        public ITextStream OpenAsTextStream(IOMode mode, Tristate format = Tristate.UseDefault)
        {
            return new MyTextStream(Path, mode);
        }
    }
    internal sealed class MyTextStream : ITextStream
    {
        private StreamReader reader;
        private StreamWriter writer;
        private readonly IOMode _mode;

        public MyTextStream(string path, IOMode mode)
        {
            _mode = mode;
            switch (mode)
            {
                case IOMode.ForReading:
                    reader = new StreamReader(path);
                    break;

                case IOMode.ForWriting:
                    writer = new StreamWriter(path, append: false);
                    break;

                case IOMode.ForAppending:
                    writer = new StreamWriter(path, append: true);
                    break;
            }
        }

        public string Read(int count)
        {
            char[] buf = new char[count];
            int r = reader.Read(buf, 0, count);
            return new string(buf, 0, r);
        }

        public string ReadAll() => reader.ReadToEnd();

        public string ReadLine() => reader.ReadLine();

        public void Write(string text) => writer.Write(text);

        public void WriteLine(string text = "") => writer.WriteLine(text);

        public void WriteBlankLines(int lines)
        {
            for (int i = 0; i < lines; i++)
                writer.WriteLine();
        }

        public void Close() => Dispose();

        public void Dispose()
        {
            writer?.Dispose();
            reader?.Dispose();
        }

        public void Skip(int count)
        {
            Read(count);
        }

        public void SkipLine()
        {
            ReadLine();
        }

        [DispId(10002)]
        public bool AtEndOfStream
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(10002)]
            get
            {
                //if (_mode == IOMode.ForReading)
                return reader.EndOfStream;
                //return true;
            }
        }

        public bool AtEndOfLine => throw new NotImplementedException();
    }
    internal enum IOMode
    {
        ForReading = 1,
        ForWriting = 2,
        ForAppending = 8
    }

    internal enum Tristate
    {
        False = 0,
        True = -1,
        UseDefault = -2
    }

    internal enum FileAttribute
    {
        Normal = 0,
        ReadOnly = 1,
        Hidden = 2,
        System = 4,
        Directory = 16,
        Archive = 32
    }

    //[ComImport]
    //[Guid("0D43FE01-F093-11CF-8940-00A0C9054228")] // public class ID
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IFileSystem
    {
        object GetDrive(string path);
        IFolder GetFolder(string path);
        IFile GetFile(string path);

        IFolder CreateFolder(string path);
        ITextStream CreateTextFile(string path, bool overwrite = false, bool unicode = false);
        ITextStream OpenTextFile(string path, IOMode mode, bool create = false, Tristate format = Tristate.UseDefault);

        bool FileExists(string path);
        bool FolderExists(string path);
        bool DriveExists(string path);
    }

    //[ComImport]
    //[Guid("0D43FE05-F093-11CF-8940-00A0C9054228")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IFolder
    {
        string Path { get; }
        string Name { get; }
        long Size { get; }

        IReadOnlyCollection<IFile> Files { get; }
        IReadOnlyCollection<IFolder> SubFolders { get; }

        void Delete(bool force);
        void Move(string newPath);
        void Copy(string newPath, bool overwrite);
    }

    //[ComImport]
    //[Guid("0D43FE03-F093-11CF-8940-00A0C9054228")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IFile
    {
        string Path { get; }
        string Name { get; }
        long Size { get; }
        string Type { get; }

        void Delete(bool force);
        void Move(string newPath);
        void Copy(string newPath, bool overwrite);

        ITextStream OpenAsTextStream(IOMode mode, Tristate format = Tristate.UseDefault);
    }

    //[ComImport]
    //[Guid("0D43FE08-F093-11CF-8940-00A0C9054228")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface ITextStream
    {
        string Read(int count);
        string ReadAll();
        string ReadLine();

        void Write(string text);
        void WriteLine(string text = "");
        void WriteBlankLines(int lines);

        void Skip(int count);
        void SkipLine();

        bool AtEndOfLine { get; }
        bool AtEndOfStream { get; }

        void Close();
    }
}