// Moved from the test project; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts and the names / values mirror the scrrun.dll type library
#pragma warning disable CA1859, CA1008, CA1027, CA1711
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in Scripting.FileSystemObject (scrrun.dll). All file system access goes through the host's IHostFileSystemHostService; if the host
    /// has not registered one, the file system of the process is used (DefaultHostFileSystemHostService, System.IO) - a host that must restrict
    /// what scripts can access registers its own service.
    /// </summary>
    [SourceClassName("FileSystemObject")] // TypeName(CreateObject("Scripting.FileSystemObject"))
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonFileSystemObject : ScriptonIReflectOnClrType, IFileSystem
    {
        public const string ProgId = "Scripting.FileSystemObject";

        private readonly IHostFileSystemHostService _hostFileSystem;

        public ScriptonFileSystemObject(IRuntimeHost runtimeHost) : base(typeof(ScriptonFileSystemObject))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _hostFileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
        }

        // --- Path functions (they do not access the file system) ---

        public string BuildPath(string Path, string Name)
        {
            if (string.IsNullOrEmpty(Path))
                return Name ?? "";
            if (string.IsNullOrEmpty(Name))
                return Path;
            bool pathEndsWithSeparator = (Path[Path.Length - 1] == '\\') || (Path[Path.Length - 1] == '/');
            bool nameStartsWithSeparator = (Name[0] == '\\') || (Name[0] == '/');
            if (pathEndsWithSeparator && nameStartsWithSeparator)
                return Path + Name.Substring(1);
            if (pathEndsWithSeparator || nameStartsWithSeparator || (Path[Path.Length - 1] == ':'))
                return Path + Name;
            // Keep the style of the path (a Windows path "C:\a" stays a Windows path on Linux too), otherwise the platform's separator
            char separator = (Path.IndexOf('\\') >= 0) ? '\\' : (Path.IndexOf('/') >= 0) ? '/' : System.IO.Path.DirectorySeparatorChar;
            return Path + separator + Name;
        }

        // The path functions work on the text only and, as with scrrun, treat both '\' and '/' as separators on every platform (System.IO.Path
        // only knows '/' on Linux)

        public string GetFileName(string Path) => FileNameOf(Path);

        /// <summary>
        /// The file name of the path (GetFileName) - '\' and '/' are separators on every platform
        /// </summary>
        internal static string FileNameOf(string path)
        {
            string trimmed = TrimTrailingSeparators(path);
            if (IsRootOnly(trimmed))
                return "";
            return trimmed.Substring(LastSeparatorIndex(trimmed) + 1);
        }

        public string GetBaseName(string Path)
        {
            string fileName = GetFileName(Path);
            int dot = fileName.LastIndexOf('.');
            return (dot < 0) ? fileName : fileName.Substring(0, dot);
        }

        public string GetExtensionName(string Path)
        {
            string fileName = GetFileName(Path);
            int dot = fileName.LastIndexOf('.');
            return (dot < 0) ? "" : fileName.Substring(dot + 1);
        }

        public string GetParentFolderName(string Path) => ParentFolderNameOf(Path);

        /// <summary>
        /// The folder of the path (GetParentFolderName) - '\' and '/' are separators on every platform
        /// </summary>
        internal static string ParentFolderNameOf(string path)
        {
            string trimmed = TrimTrailingSeparators(path);
            if (IsRootOnly(trimmed))
                return "";
            int separator = LastSeparatorIndex(trimmed);
            if (separator < 0)
                return "";
            string parent = trimmed.Substring(0, separator);
            return ((parent.Length == 2) && (parent[1] == ':')) ? parent + trimmed[separator] : parent; // "C:\a" -> "C:\"
        }

        public string GetDriveName(string Path)
        {
            string path = Path ?? "";
            return ((path.Length >= 2) && (path[1] == ':')) ? path.Substring(0, 2) : "";
        }

        private static string TrimTrailingSeparators(string path)
        {
            return (path ?? "").TrimEnd('\\', '/'); // "C:\a\dir\" -> "C:\a\dir", "C:\" -> "C:"
        }

        private static bool IsRootOnly(string trimmedPath) => (trimmedPath.Length == 0) || ((trimmedPath.Length == 2) && (trimmedPath[1] == ':'));

        private static int LastSeparatorIndex(string path) => Math.Max(path.LastIndexOf('\\'), path.LastIndexOf('/'));

        public string GetAbsolutePathName(string Path) => System.IO.Path.GetFullPath(string.IsNullOrEmpty(Path) ? "." : Path);

        /// <summary>
        /// A random file name like "radC2A1B.tmp" (as scrrun returns), the file is not created
        /// </summary>
        public string GetTempName() => "rad" + Guid.NewGuid().ToString("N").Substring(0, 5).ToUpperInvariant() + ".tmp";

        // --- File and folder operations ---

        public void DeleteFile(string FileSpec, [Optional, DefaultParameterValue(false)] bool Force)
        {
            if (!_hostFileSystem.FileExists(FileSpec))
                throw FileNotFound(FileSpec);
            _hostFileSystem.DeleteFile(FileSpec);
        }

        public void DeleteFolder(string FolderSpec, [Optional, DefaultParameterValue(false)] bool Force)
        {
            if (!_hostFileSystem.DirectoryExists(FolderSpec))
                throw PathNotFound(FolderSpec);
            _hostFileSystem.DeleteDirectory(FolderSpec, recursive: true);
        }

        public void CopyFile(string Source, string Destination, [Optional, DefaultParameterValue(true)] bool OverWriteFiles)
        {
            _hostFileSystem.CopyFile(Source, IsFolderTarget(Destination) ? BuildPath(Destination, GetFileName(Source)) : Destination, OverWriteFiles);
        }

        public void MoveFile(string Source, string Destination)
        {
            _hostFileSystem.MoveFile(Source, IsFolderTarget(Destination) ? BuildPath(Destination, GetFileName(Source)) : Destination);
        }

        public void CopyFolder(string Source, string Destination, [Optional, DefaultParameterValue(true)] bool OverWriteFiles)
        {
            _hostFileSystem.CopyDirectory(Source, Destination, OverWriteFiles);
        }

        public void MoveFolder(string Source, string Destination)
        {
            _hostFileSystem.MoveDirectory(Source, Destination);
        }

        /// <summary>
        /// As with scrrun, a destination that ends with a path separator (or is an existing folder) is the folder to copy / move the file into
        /// </summary>
        private bool IsFolderTarget(string destination)
        {
            return !string.IsNullOrEmpty(destination)
                && ((destination[destination.Length - 1] == '\\') || (destination[destination.Length - 1] == '/') || _hostFileSystem.DirectoryExists(destination));
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
            throw new NotSupportedException($"FileSystemObject.GetDrive('{path}'): Drive objects are not supported (DriveExists and GetDriveName are)");
        }
        public IFile GetFile(string path)
        {
            if (!_hostFileSystem.FileExists(path))
                throw FileNotFound(path);
            HostFileSystemFileInfo nfo = _hostFileSystem.GetFileInfo(path);
            if (!nfo.Exists)
            {
                throw FileNotFound(path);
            }

            return new ScriptonFile(_hostFileSystem, nfo);
        }

        public IFolder GetFolder(string path)
        {
            if (!_hostFileSystem.DirectoryExists(path))
                throw PathNotFound(path);
            HostFileSystemDirectoryInfo nfo = _hostFileSystem.GetDirectoryInfo(path);
            if (!nfo.Exists)
            {
                throw PathNotFound(path);
            }

            return new ScriptonFolder(_hostFileSystem, nfo);
        }

        public IFolder CreateFolder(string path)
        {
            _ = nameof(ScriptonFolder);
            HostFileSystemDirectoryInfo nfo = _hostFileSystem.CreateDirectory(path);
            return new ScriptonFolder(_hostFileSystem, nfo);
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

            if (!Enum.IsDefined(typeof(IOMode), mode))
                throw new ArgumentException($"Undefined mode:{mode}", nameof(mode));

            if (!create && !_hostFileSystem.FileExists(path))
                throw FileNotFound(path); // error 53, as scrrun
            if (mode == IOMode.ForReading)
            {
                return new ScriptonFileTextStream(path, _hostFileSystem.OpenTextFileRead(path));
            }
            else
            {
                return new ScriptonFileTextStream(path, _hostFileSystem.OpenTextFileWrite(path, createIfNotExists: create, overwriteIfExists: true, mode == IOMode.ForAppending), mode, unicode: false);
            }
        }

        /// <summary>
        /// CreateTextFile(FileName, [Overwrite = True], [Unicode = False]) - creates the file; an existing file is replaced unless Overwrite is False,
        /// in which case error 58 (File already exists) is raised (verified with Windows scrrun)
        /// </summary>
        public ITextStream CreateTextFile(string path, bool overwrite = true, bool unicode = false)
        {
            if (!overwrite && _hostFileSystem.FileExists(path))
                throw new CustomException(58, ErrorSource, "File already exists");
            return new ScriptonFileTextStream(path, _hostFileSystem.OpenTextFileWrite(path, createIfNotExists: true, overwriteIfExists: overwrite, append: false), IOMode.ForWriting, unicode);
        }

        private const string ErrorSource = "Microsoft VBScript runtime error";

        /// <summary>
        /// Error 53, File not found (as scrrun raises for a missing file)
        /// </summary>
        internal static CustomException FileNotFound(string path) => new CustomException(53, ErrorSource, "File not found: " + path);

        /// <summary>
        /// Error 76, Path not found (as scrrun raises for a missing folder)
        /// </summary>
        internal static CustomException PathNotFound(string path) => new CustomException(76, ErrorSource, "Path not found: " + path);

        [DispId(10014)]
        public IFolder GetSpecialFolder([In] object SpecialFolder)
        {
            SpecialFolderConst specialFolderX = (SpecialFolderConst)Enum.ToObject(typeof(SpecialFolderConst), SpecialFolder);
            if (specialFolderX == SpecialFolderConst.TemporaryFolder)
            {
                HostFileSystemDirectoryInfo nfo = _hostFileSystem.GetSpecialFolderTemp();
                return new ScriptonFolder(_hostFileSystem, nfo);
            }
            throw new NotSupportedException($"FileSystemObject.GetSpecialFolder({SpecialFolder}): only the TemporaryFolder (2) is supported");
        }
    }

    [SourceClassName("Folder")]
    [ComVisible(true)]
    [DefaultMember("Path")] // +[DispId(0)] +[IsDefault]
    internal sealed class ScriptonFolder : IFolder
    {
        private readonly IHostFileSystemHostService _hostFileSystem;
        private readonly HostFileSystemDirectoryInfo _info;
        [DispId(0)] [IsDefault] public string Path => _info.Path;
        public string Name => _info.Name;

        public ScriptonFolder(IHostFileSystemHostService hostFileSystemService, HostFileSystemDirectoryInfo info)
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

        /// <summary>
        /// The files of the folder - a Files collection (Count, Item(name), For Each), as with scrrun
        /// </summary>
        public IReadOnlyCollection<IFile> Files => new ScriptonFiles(_hostFileSystem.GetFiles(Path).Select(nfo => (IFile)new ScriptonFile(_hostFileSystem, nfo)).ToList());

        /// <summary>
        /// The folders of the folder - a Folders collection (Count, Item(name), For Each), as with scrrun
        /// </summary>
        public IReadOnlyCollection<IFolder> SubFolders => new ScriptonFolders(_hostFileSystem.GetDirectories(Path).Select(nfo => (IFolder)new ScriptonFolder(_hostFileSystem, nfo)).ToList());

        public void Delete(bool force) => _hostFileSystem.DeleteDirectory(Path, force);

        public void Move(string newPath) => _hostFileSystem.MoveDirectory(Path, newPath);

        public void Copy(string newPath, bool overwrite)
        {
            _hostFileSystem.CopyDirectory(Path, newPath, overwrite);
        }
    }

    [SourceClassName("File")]
    [ComVisible(true)]
    [DefaultMember("Path")]
    internal sealed class ScriptonFile : IFile
    {
        private readonly IHostFileSystemHostService _hostFileSystem;
        private readonly HostFileSystemFileInfo _nfo;
        [DispId(0)] [IsDefault] public string Path => _nfo.Path;
        public string Name => _nfo.Name;

        /// <summary>
        /// The size in bytes - a Long, or a Double for files of 2 GB and more (as with scrrun)
        /// </summary>
        public object Size => (_nfo.Size <= int.MaxValue) ? (object)(int)_nfo.Size : (double)_nfo.Size;

        public DateTime DateCreated => _nfo.DateCreated;

        public DateTime DateLastModified => _nfo.DateLastModified;

        public DateTime DateLastAccessed => _nfo.DateLastAccessed;

        public string ParentFolder => ScriptonFileSystemObject.ParentFolderNameOf(_nfo.Path);

        public ScriptonFile(IHostFileSystemHostService hostFileSystem, HostFileSystemFileInfo nfo)
        {
            _hostFileSystem = hostFileSystem ?? throw new ArgumentNullException(nameof(hostFileSystem));
            _nfo = nfo ?? throw new ArgumentNullException(nameof(nfo));
        }

        public void Delete(bool force)
        {
            _hostFileSystem.DeleteFile(_nfo.Path);
        }
        public void Delete()
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
            if (!Enum.IsDefined(typeof(IOMode), mode))
                throw new ArgumentException($"Undefined mode:{mode}", nameof(mode));

            if (mode == IOMode.ForReading)
            {
                return new ScriptonFileTextStream(Path, _hostFileSystem.OpenTextFileRead(Path));
            }
            else
            {
                return new ScriptonFileTextStream(Path, _hostFileSystem.OpenTextFileWrite(Path, createIfNotExists: true, overwriteIfExists: mode == IOMode.ForAppending, append: mode == IOMode.ForAppending), mode, unicode: false);
            }
        }
    }

    [SourceClassName("TextStream")]
    [ComVisible(true)]
    [DebuggerDisplay("({_mode}) Path:{_path}")]
    internal sealed class ScriptonFileTextStream : ITextStream, IDisposable
    {
        private System.IO.StreamReader _reader; // nullable
        private System.IO.StreamWriter _writer; // nullable
        private System.IO.Stream _writeStream; // nullable
        private readonly string _path;
        private readonly IOMode _mode;

        public ScriptonFileTextStream(string path, System.IO.StreamReader reader)
        {
            _mode = IOMode.ForReading;
            _path = path ?? throw new ArgumentNullException(nameof(path));
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        }
        public ScriptonFileTextStream(string path, System.IO.Stream writeStream, IOMode mode, bool unicode)
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

        public bool AtEndOfLine
        {
            get
            {
                int next = reader.Peek();
                return (next == -1) || (next == 13) || (next == 10); // end of stream, CR or LF
            }
        }
    }
    public enum IOMode
    {
        ForReading = 1,
        ForWriting = 2,
        ForAppending = 8
    }

    public enum Tristate
    {
        False = 0,
        True = -1,
        UseDefault = -2
    }

    public enum FileAttribute
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
    public interface IFileSystem
    {
        string BuildPath(string Path, string Name);
        string GetFileName(string Path);
        string GetBaseName(string Path);
        string GetExtensionName(string Path);
        string GetParentFolderName(string Path);
        string GetDriveName(string Path);
        string GetAbsolutePathName(string Path);
        string GetTempName();
        void DeleteFile(string FileSpec, bool Force = false);
        void DeleteFolder(string FolderSpec, bool Force = false);
        void CopyFile(string Source, string Destination, bool OverWriteFiles = true);
        void MoveFile(string Source, string Destination);
        void CopyFolder(string Source, string Destination, bool OverWriteFiles = true);
        void MoveFolder(string Source, string Destination);

        object GetDrive(string path);
        IFolder GetFolder(string path);
        IFile GetFile(string path);

        IFolder CreateFolder(string path);
        ITextStream CreateTextFile(string path, bool overwrite = true, bool unicode = false); // Overwrite defaults to True, as scrrun
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
    public interface IFolder
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

    /// <summary>
    /// Folder.Files / Folder.SubFolders: Count, Item(name or path) and For Each (TypeName "Files" / "Folders", as with scrrun)
    /// </summary>
    [ComVisible(true)]
    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault]
    internal abstract class ScriptonFileSystemCollection<T> : IReadOnlyCollection<T> where T : class
    {
        private readonly List<T> _items;
        private readonly Func<T, string> _name;
        private readonly Func<T, string> _path;

        protected ScriptonFileSystemCollection(List<T> items, Func<T, string> name, Func<T, string> path)
        {
            _items = items;
            _name = name;
            _path = path;
        }

        public int Count => _items.Count;

        /// <summary>
        /// Item(name) - the file or folder with the name (or the full path); error 53 (File not found) / 76 (Path not found) if there is none
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public T Item(object name)
        {
            string key = Convert.ToString(name, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            return _items.FirstOrDefault(i => string.Equals(_name(i), key, StringComparison.OrdinalIgnoreCase) || string.Equals(_path(i), key, StringComparison.OrdinalIgnoreCase))
                ?? throw ((typeof(T) == typeof(IFile)) ? ScriptonFileSystemObject.FileNotFound(key) : ScriptonFileSystemObject.PathNotFound(key));
        }

        public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [SourceClassName("Files")]
    [ComVisible(true)]
    [DefaultMember("Item")]
    internal sealed class ScriptonFiles : ScriptonFileSystemCollection<IFile>
    {
        public ScriptonFiles(List<IFile> files) : base(files, f => f.Name, f => f.Path)
        {
        }
    }

    [SourceClassName("Folders")]
    [ComVisible(true)]
    [DefaultMember("Item")]
    internal sealed class ScriptonFolders : ScriptonFileSystemCollection<IFolder>
    {
        public ScriptonFolders(List<IFolder> folders) : base(folders, f => f.Name, f => f.Path)
        {
        }
    }

    //[ComImport]
    //[Guid("0D43FE03-F093-11CF-8940-00A0C9054228")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IFile
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
    public interface ITextStream
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