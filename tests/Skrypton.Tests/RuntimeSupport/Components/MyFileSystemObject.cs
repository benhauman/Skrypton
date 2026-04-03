using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.Tests.RuntimeSupport.Components;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.FileSystemSupport
{
    [SourceClassName("Dictionary")] // for TYPENAME(CreateObject("Scripting.Dictionary"))
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    //[DefaultMember("Item")]
    internal sealed class MyFileSystemObject : IReflectOnClrType, IFileSystem
    {
        /* COM:"Microsoft Scripting Runtime" => 'Interop.Scripting.dll' => C:\Windows\System32\scrrun.dll */
        // D:\projects.ToDelete\ConsoleApp2\bin\Debug\net10.0\Interop.Scripting.dll

        private readonly IHostFileSystemHostService _hostFileSystem;
        public MyFileSystemObject(IHostFileSystemHostService hostFileSystemService)
        {
            _hostFileSystem = hostFileSystemService ?? throw new ArgumentNullException(nameof(hostFileSystemService));
        }
        internal static MyFileSystemObject Create(IServiceProvider hostServices)
        {
            return new MyFileSystemObject(hostServices.GetRequiredService<IHostFileSystemHostService>());
        }

        public bool FileExists(string path)
        {
            return _hostFileSystem.FileExists(path);
        }
        public bool FolderExists(string path)
        {
            return _hostFileSystem.DirectoryExists(path);
        }
        public bool DriveExists(string path)
        {
            return _hostFileSystem.DriveExists(path);
        }

        public object GetDrive(string path)
        {
            //string root = Path.GetPathRoot(path);
            //if (string.IsNullOrEmpty(root))
            //    return false;

            //// Normalize like "C:\"
            //root = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

            //var nfo = DriveInfo.GetDrives()
            //    .Any(d => string.Equals(
            //        d.Name,
            //        root,
            //        StringComparison.OrdinalIgnoreCase));

            throw new NotImplementedException(path);
        }
        public IFile GetFile(string path)
        {
            //throw new NotImplementedException(path);
            if (!_hostFileSystem.FileExists(path))
                throw new System.IO.FileNotFoundException(path);
            HostFileSystemFileInfo nfo = _hostFileSystem.GetFileInfo(path);
            if (!nfo.Exists)
            {
                throw new System.IO.FileNotFoundException(path);
            }

            return new MyFile(_hostFileSystem, nfo);
        }

        public IFolder GetFolder(string path)
        {
            if (!_hostFileSystem.DirectoryExists(path))
                throw new System.IO.DirectoryNotFoundException(path);
            HostFileSystemDirectoryInfo nfo = _hostFileSystem.GetDirectoryInfo(path);
            if (!nfo.Exists)
            {
                throw new System.IO.DirectoryNotFoundException(path);
            }

            return new MyFolder(_hostFileSystem, nfo);
        }

        public IFolder CreateFolder(string path)
        {
            _ = nameof(MyFolder);
            HostFileSystemDirectoryInfo nfo = _hostFileSystem.CreateDirectory(path);
            return new MyFolder(_hostFileSystem, nfo);
        }

        public ITextStream OpenTextFile(string path, short mode) => OpenTextFileCore(path, (IOMode)mode, create:false, format: Tristate.UseDefault);
        public ITextStream OpenTextFile(string path, IOMode mode) => OpenTextFileCore(path, mode, create: false, format: Tristate.UseDefault);
        public ITextStream OpenTextFile(string path, short mode, bool create) => OpenTextFileCore(path, (IOMode)mode, create, format: Tristate.UseDefault);
        public ITextStream OpenTextFile(string path, IOMode mode, bool create) => OpenTextFileCore(path, mode, create, format: Tristate.UseDefault);
        public ITextStream OpenTextFile(string path, short mode, bool create, short format) => OpenTextFileCore(path, (IOMode)mode, create, (Tristate)format);
        public ITextStream OpenTextFile(string path, IOMode mode, bool create, Tristate format) => OpenTextFileCore(path, mode, create, format);
        private ITextStream OpenTextFileCore(string path, IOMode mode, bool create, Tristate format)
        {
            //if (mode == null)
            //    throw new ArgumentNullException($"Argument 'mode' is required", nameof(mode));

            if (!Enum.IsDefined<IOMode>(mode))
                throw new ArgumentException($"Undefined mode:{mode}", nameof(mode));

            if (mode == IOMode.ForReading)
            {
                return new MyTextStream(path, _hostFileSystem.OpenTextFileRead(path));
            }
            else
            {
                return new MyTextStream(path, _hostFileSystem.OpenTextFileWrite(path, createIfNotExists: create, overwriteIfExists: true, mode == IOMode.ForAppending), mode, unicode: false);
            }
        }

        public ITextStream CreateTextFile(string path, bool overwrite = true, bool unicode = false)
        {
            return new MyTextStream(path, _hostFileSystem.OpenTextFileWrite(path, createIfNotExists: overwrite, overwriteIfExists: true, false), IOMode.ForWriting, unicode);
        }

        [DispId(10014)]
        public IFolder GetSpecialFolder([In] object SpecialFolder)
        {
            SpecialFolderConst specialFolderX = (SpecialFolderConst)Enum.ToObject(typeof(SpecialFolderConst), SpecialFolder);
            if (specialFolderX == SpecialFolderConst.TemporaryFolder)
            {
                HostFileSystemDirectoryInfo nfo = _hostFileSystem.GetSpecialFolderTemp();
                return new MyFolder(_hostFileSystem, nfo);
            }
            throw new NotImplementedException($"SpecialFolder:{SpecialFolder}");
        }
    }

    [DefaultMember("Path")] // +[DispId(0)] +[IsDefault]
    internal sealed class MyFolder : IFolder
    {
        private readonly IHostFileSystemHostService _hostFileSystem;
        private readonly HostFileSystemDirectoryInfo _info;
        [DispId(0)] [IsDefault] public string Path => _info.Path;
        public string Name => _info.Name;

        public MyFolder(IHostFileSystemHostService hostFileSystemService, HostFileSystemDirectoryInfo info)
        {
            _hostFileSystem = hostFileSystemService ?? throw new ArgumentNullException(nameof(hostFileSystemService));
            _info = info ?? throw new ArgumentNullException(nameof(info));
        }

        //public long Size
        //{
        //    get
        //    {
        //        long sum = 0;
        //        foreach (var f in Directory.GetFiles(Path, "*", SearchOption.AllDirectories)) // works on Windows, Linux, and macOS when using .NET Standard or modern .NET.
        //            sum += new FileInfo(f).Length;
        //        return sum;
        //    }
        //}

        public IReadOnlyCollection<IFile> Files => _hostFileSystem.GetFiles(Path).Select(nfo => new MyFile(_hostFileSystem, nfo)).ToArray();
        public IReadOnlyCollection<IFolder> SubFolders => _hostFileSystem.GetDirectories(Path).Select(nfo => new MyFolder(_hostFileSystem, nfo)).ToArray();

        public void Delete(bool force) => _hostFileSystem.DeleteDirectory(Path, force);

        public void Move(string newPath) => _hostFileSystem.MoveDirectory(Path, newPath);

        public void Copy(string newPath, bool overwrite)
        {
            _hostFileSystem.CopyDirectory(Path, newPath, overwrite);
        }
    }

    internal sealed class MyFile : IFile
    {
        private readonly IHostFileSystemHostService _hostFileSystem;
        private readonly HostFileSystemFileInfo _nfo;
        public string Path => _nfo.Path;
        public string Name => _nfo.Name;
        //public long Size => new FileInfo(Path).Length;
        //public string Type => "File";

        public MyFile(IHostFileSystemHostService hostFileSystem, HostFileSystemFileInfo nfo)
        {
            _hostFileSystem = hostFileSystem ?? throw new ArgumentNullException(nameof(hostFileSystem));
            _nfo = nfo ?? throw new ArgumentNullException(nameof(nfo));
        }

        public void Delete(bool force)
        {
            _hostFileSystem.DeleteFile(_nfo.Path);
        }

        public void Move(string newPath)
        {
            _hostFileSystem.MoveFile(Path, newPath);
        }

        public void Copy(string newPath, bool overwrite)
        {
            _hostFileSystem.CopyFile(Path, newPath, overwrite);
        }

        public ITextStream OpenAsTextStream(IOMode mode, Tristate format = Tristate.UseDefault)
        {
            if (!Enum.IsDefined<IOMode>(mode))
                throw new ArgumentException($"Undefined mode:{mode}", nameof(mode));

            if (mode == IOMode.ForReading)
            {
                return new MyTextStream(Path, _hostFileSystem.OpenTextFileRead(Path));
            }
            else
            {
                return new MyTextStream(Path, _hostFileSystem.OpenTextFileWrite(Path, createIfNotExists: true, overwriteIfExists: mode == IOMode.ForAppending, append: mode == IOMode.ForAppending), mode, unicode: false);
            }
        }
    }

    [DebuggerDisplay("({_mode}) Path:{_path}")]
    internal sealed class MyTextStream : ITextStream, IDisposable
    {
        private System.IO.StreamReader _reader; // nullable
        private System.IO.StreamWriter _writer; // nullable
        private System.IO.Stream _writeStream; // nullable
        private readonly string _path;
        private readonly IOMode _mode;

        public MyTextStream(string path, System.IO.StreamReader reader)
        {
            _mode = IOMode.ForReading;
            _path = path ?? throw new ArgumentNullException(nameof(path));
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        }
        public MyTextStream(string path, System.IO.Stream writeStream, IOMode mode, bool unicode)
        {
            _mode = mode;
            _path = path ?? throw new ArgumentNullException(nameof(path));
            _writeStream = writeStream ?? throw new ArgumentNullException(nameof(writeStream));
            _writer = new System.IO.StreamWriter(_writeStream, unicode ? System.Text.Encoding.Unicode : System.Text.Encoding.ASCII);
        }
        private System.IO.StreamReader reader => _reader ?? throw new InvalidOperationException($"Reader not set. Path:{_path}");
        private System.IO.StreamWriter writer => _writer ?? throw new InvalidOperationException($"Writer not set. ({_mode}) Path:{_path}");

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
            _writer?.Dispose();
            _writer = null;
            _writeStream?.Dispose();
            _writeStream = null;
            _reader?.Dispose();
            _reader = null;
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

        [DispId(10014)]
        IFolder GetSpecialFolder([In]object SpecialFolder);
    }

    public enum SpecialFolderConst
    {
        WindowsFolder,
        SystemFolder,
        TemporaryFolder
    }


    //[ComImport]
    //[Guid("0D43FE05-F093-11CF-8940-00A0C9054228")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IFolder
    {
        string Path { get; }
        string Name { get; }
        //long Size { get; }

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
        //long Size { get; }
        //string Type { get; }

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