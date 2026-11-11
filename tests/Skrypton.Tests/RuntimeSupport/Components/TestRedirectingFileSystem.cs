#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// An IHostFileSystemHostService on the real file system where some paths are redirected - eg. the path of a customer template that a script
    /// opens ("C:\helpline\Dokumente\..docx") to a test fixture written to a temporary file
    /// </summary>
    internal sealed class TestRedirectingFileSystem : IHostFileSystemHostService
    {
        private readonly Dictionary<string, string> _redirections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public TestRedirectingFileSystem Redirect(string path, string target)
        {
            _redirections[path] = target;
            return this;
        }

        private string R(string path) => _redirections.TryGetValue(path ?? "", out string? target) ? target : path!;

        public bool FileExists(string path) => File.Exists(R(path));
        public bool DirectoryExists(string path) => Directory.Exists(R(path));
        public StreamReader OpenTextFileRead(string path) => new StreamReader(new FileStream(R(path), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete));

        public Stream OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append)
        {
            string target = R(path);
            FileMode mode = append ? FileMode.Append : (overwriteIfExists ? FileMode.Create : FileMode.CreateNew);
            if (!createIfNotExists && !File.Exists(target))
                throw new FileNotFoundException(target);
            return new FileStream(target, mode, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
        }

        public HostFileSystemDirectoryInfo CreateDirectory(string path) => ToInfo(Directory.CreateDirectory(R(path)));
        public void DeleteFile(string path) => File.Delete(R(path));
        public void DeleteDirectory(string path, bool recursive) => Directory.Delete(R(path), recursive);
        public void MoveFile(string src, string dst) => File.Move(R(src), R(dst));
        public void MoveDirectory(string src, string dst) => Directory.Move(R(src), R(dst));
        public void CopyFile(string src, string dst, bool overwrite) => File.Copy(R(src), R(dst), overwrite);
        public IEnumerable<HostFileSystemFileInfo> GetFiles(string directory) => new DirectoryInfo(R(directory)).GetFiles().Select(f => new HostFileSystemFileInfo(f.FullName, f.Name, true));
        public IEnumerable<HostFileSystemDirectoryInfo> GetDirectories(string directory) => new DirectoryInfo(R(directory)).GetDirectories().Select(ToInfo);
        public bool DriveExists(string path) => Directory.Exists(Path.GetPathRoot(R(path)));
        public HostFileSystemDirectoryInfo GetDirectoryInfo(string path) => ToInfo(new DirectoryInfo(R(path)));

        public HostFileSystemFileInfo GetFileInfo(string path)
        {
            var info = new FileInfo(R(path));
            return new HostFileSystemFileInfo(info.FullName, info.Name, info.Exists);
        }

        public void CopyDirectory(string sourcePath, string newPath, bool overwrite) => throw new NotSupportedException("CopyDirectory");
        public HostFileSystemDirectoryInfo GetSpecialFolderTemp() => ToInfo(new DirectoryInfo(Path.GetTempPath()));

        private static HostFileSystemDirectoryInfo ToInfo(DirectoryInfo info) => new HostFileSystemDirectoryInfo(info.FullName, info.Name, info.Exists);
    }
}
