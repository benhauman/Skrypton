using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The file system of the process (System.IO) - used by the built-in Scripting.FileSystemObject when the host has not registered an
    /// IHostFileSystemHostService of its own (moved from the test project, where it was WindowsFileSystem; it is not Windows specific)
    /// </summary>
    internal sealed class DefaultHostFileSystemHostService : HostFileSystemHostServiceBase, IHostFileSystemHostService
    {
        public static readonly DefaultHostFileSystemHostService Instance = new DefaultHostFileSystemHostService();

        private DefaultHostFileSystemHostService()
        {
        }

        public bool FileExists(string path) => File.Exists(path);
        public bool DirectoryExists(string path) => Directory.Exists(path);

        // FileShare.Delete: as with scrrun, a file that a script still has open (a TextStream that is not closed yet - VBScript releases objects
        // immediately, .NET only when collected) can be deleted or replaced
        public StreamReader OpenTextFileRead(string path) => new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete));
        public Stream OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append)// => new StreamWriter(path, append, unicode ? Encoding.Unicode : null);
        {
            if (path == null) throw new ArgumentNullException(nameof(path));

            if (File.Exists(path))
            {
                if (append)
                {
                    return new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
                }
                else
                {
                    if (overwriteIfExists)
                    {
                        return new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
                    }
                    else
                    {
                        throw new System.IO.IOException("File overwrite is forbidden. file: " + path);
                    }
                }
            }
            else
            {
                if (createIfNotExists)
                {
                    return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
                }
                else
                {
                    throw new System.IO.IOException("File not found. file: " + path);
                }
            }
        }

        public HostFileSystemDirectoryInfo CreateDirectory(string path) => new HostFileSystemDirectoryInfo(path, Directory.CreateDirectory(path).Name, true);

        public void DeleteFile(string path) => File.Delete(path);
        public void DeleteDirectory(string path, bool recursive) =>
            Directory.Delete(path, recursive);

        public void MoveFile(string src, string dst) => File.Move(src, dst);
        public void MoveDirectory(string src, string dst) => Directory.Move(src, dst);

        public void CopyFile(string src, string dst, bool overwrite) =>
            File.Copy(src, dst, overwrite);

        public IEnumerable<HostFileSystemFileInfo> GetFiles(string directory) => Directory.GetFiles(directory).Select(x => new FileInfo(x)).Select(ToFileInfo).ToArray();

        public IEnumerable<HostFileSystemDirectoryInfo> GetDirectories(string directory)
        {
            return Directory.GetDirectories(directory).Select(fullName => new HostFileSystemDirectoryInfo(fullName, new DirectoryInfo(fullName).Name, true)).ToArray();
        }

        public bool DriveExists(string path)
        {
            string root = Path.GetPathRoot(path);
            if (string.IsNullOrEmpty(root)) return false;

            root = root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

            return DriveInfo.GetDrives()
                .Any(d => string.Equals(d.Name, root, StringComparison.OrdinalIgnoreCase));
        }

        public HostFileSystemDirectoryInfo GetDirectoryInfo(string path)
        {
            DirectoryInfo nfo = new DirectoryInfo(path);
            if (!nfo.Exists)
                throw new System.IO.DirectoryNotFoundException(path);
            return new HostFileSystemDirectoryInfo(path, nfo.Name, nfo.Exists);
        }
        private static HostFileSystemFileInfo ToFileInfo(FileInfo info)
        {
            return info.Exists
                ? new HostFileSystemFileInfo(info.FullName, info.Name, true, info.Length, info.CreationTime, info.LastWriteTime, info.LastAccessTime)
                : new HostFileSystemFileInfo(info.FullName, info.Name, false);
        }

        public HostFileSystemFileInfo GetFileInfo(string path)
        {
            var nfo = new FileInfo(path);
            if (!nfo.Exists)
                throw new System.IO.FileNotFoundException(path);
            return ToFileInfo(nfo);
        }

        public void CopyDirectory(string sourcePath, string newPath, bool overwrite)
        {
            Directory.CreateDirectory(newPath);
            foreach (var file in Directory.GetFiles(sourcePath))
            {
                var dest = System.IO.Path.Combine(newPath, System.IO.Path.GetFileName(file));
                File.Copy(file, dest, overwrite);
            }
        }

        public HostFileSystemDirectoryInfo GetSpecialFolderTemp()
        {
            string path = Path.GetTempPath();
            DirectoryInfo nfo = new DirectoryInfo(path);
            if (!nfo.Exists)
                throw new System.IO.DirectoryNotFoundException(path);
            return new HostFileSystemDirectoryInfo(path, nfo.Name, nfo.Exists);
        }
    }
}
