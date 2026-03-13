using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Skrypton.Tests.RuntimeSupport.Implementations.FileSystemSupport
{

    internal sealed class WindowsFileSystem : HostFileSystemHostServiceBase, IHostFileSystemHostService
    {
        public bool FileExists(string path) => File.Exists(path);
        public bool DirectoryExists(string path) => Directory.Exists(path);

        public StreamReader OpenTextFileRead(string path) => File.OpenText(path);
        public Stream OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append)// => new StreamWriter(path, append, unicode ? Encoding.Unicode : null);
        {
            if (path == null) throw new ArgumentNullException(nameof(path));

            if (File.Exists(path))
            {
                if (append)
                {
                    return new FileStream(path, FileMode.Append, FileAccess.Write);
                }
                else
                {
                    if (overwriteIfExists)
                    {
                        return new FileStream(path, FileMode.Create, FileAccess.Write);
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
                    return new FileStream(path, FileMode.CreateNew, FileAccess.Write);
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

        public IEnumerable<HostFileSystemFileInfo> GetFiles(string directory) => Directory.GetFiles(directory).Select(x => new FileInfo(x)).Select(x => new HostFileSystemFileInfo(x.FullName, x.Name, x.Exists)).ToArray();

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
        public HostFileSystemFileInfo GetFileInfo(string path)
        {
            var nfo = new FileInfo(path);
            if (!nfo.Exists)
                throw new System.IO.FileNotFoundException(path);
            return new HostFileSystemFileInfo(nfo.FullName, nfo.Name, nfo.Exists);
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
    }

    internal sealed class LinuxFileSystem : HostFileSystemHostServiceBase, IHostFileSystemHostService
    {
        public bool FileExists(string path) => File.Exists(path);
        public bool DirectoryExists(string path) => Directory.Exists(path);

        public StreamReader OpenTextFileRead(string path) => File.OpenText(path);
        public Stream OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append)// => new StreamWriter(path, append, unicode ? Encoding.Unicode : null);
        {
            if (path == null) throw new ArgumentNullException(nameof(path));

            if (File.Exists(path))
            {
                if (append)
                {
                    return new FileStream(path, FileMode.Append, FileAccess.Write);
                }
                else
                {
                    if (overwriteIfExists)
                    {
                        return new FileStream(path, FileMode.Create, FileAccess.Write);
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
                    return new FileStream(path, FileMode.CreateNew, FileAccess.Write);
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

        public IEnumerable<HostFileSystemFileInfo> GetFiles(string directory) => Directory.GetFiles(directory).Select(x => new FileInfo(x)).Select(x => new HostFileSystemFileInfo(x.FullName, x.Name, x.Exists)).ToArray();

        public IEnumerable<HostFileSystemDirectoryInfo> GetDirectories(string directory)
        {
            return Directory.GetDirectories(directory).Select(fullName => new HostFileSystemDirectoryInfo(fullName, new DirectoryInfo(fullName).Name, true)).ToArray();
        }

        public bool DriveExists(string path)
        {
            // Unix has no “drives” — treat the root (/) as always existing
            string root = Path.GetPathRoot(path);
            if (string.IsNullOrEmpty(root)) return false;

            // Usually root = "/"
            return Directory.Exists(root);
        }
        public HostFileSystemDirectoryInfo GetDirectoryInfo(string path)
        {
            DirectoryInfo nfo = new DirectoryInfo(path);
            if (!nfo.Exists)
                throw new System.IO.DirectoryNotFoundException(path);
            return new HostFileSystemDirectoryInfo(path, nfo.Name, nfo.Exists);
        }
        public HostFileSystemFileInfo GetFileInfo(string path)
        {
            var nfo = new FileInfo(path);
            if (!nfo.Exists)
                throw new System.IO.FileNotFoundException(path);
            return new HostFileSystemFileInfo(nfo.FullName, nfo.Name, nfo.Exists);
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
    }

    internal sealed class TestFileSystem : IHostFileSystemHostService
    {
        private readonly Dictionary<string, StringBuilder> _allfiles = new Dictionary<string, StringBuilder>(StringComparer.OrdinalIgnoreCase);

        public TestFileSystem()
        {
        }

        void IHostFileSystemHostService.CopyDirectory(string sourcePath, string newPath, bool overwrite)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.CopyFile(string src, string dst, bool overwrite)
        {
            throw new NotImplementedException();
        }

        HostFileSystemDirectoryInfo IHostFileSystemHostService.CreateDirectory(string path)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.DeleteDirectory(string path, bool recursive)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.DeleteFile(string path)
        {
            throw new NotImplementedException();
        }

        bool IHostFileSystemHostService.DirectoryExists(string path)
        {
            throw new NotImplementedException();
        }

        bool IHostFileSystemHostService.DriveExists(string path)
        {
            throw new NotImplementedException();
        }

        bool IHostFileSystemHostService.FileExists(string path)
        {
            throw new NotImplementedException();
        }

        IEnumerable<HostFileSystemDirectoryInfo> IHostFileSystemHostService.GetDirectories(string directory)
        {
            throw new NotImplementedException();
        }

        HostFileSystemDirectoryInfo IHostFileSystemHostService.GetDirectoryInfo(string path)
        {
            throw new NotImplementedException();
        }

        HostFileSystemFileInfo IHostFileSystemHostService.GetFileInfo(string path)
        {
            throw new NotImplementedException();
        }

        IEnumerable<HostFileSystemFileInfo> IHostFileSystemHostService.GetFiles(string directory)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.MoveDirectory(string src, string dst)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.MoveFile(string src, string dst)
        {
            throw new NotImplementedException();
        }

        StreamReader IHostFileSystemHostService.OpenTextFileRead(string path)
        {
            if (_allfiles.TryGetValue(path, out StringBuilder content))
            {
                byte[] buffer = Encoding.UTF8.GetBytes(content.ToString());
                return new StreamReader(new MemoryStream(buffer));
            }

            throw new NotImplementedException($"[FS].OpenTextFileWrite(path:'{path}'");
        }

        Stream IHostFileSystemHostService.OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append)
        {
            if (_allfiles.TryGetValue(path, out StringBuilder content))
            {
                return new StringStream(content);
            }

            throw new NotImplementedException($"[FS].OpenTextFileWrite(path:'{path}', createIfNotExists:{createIfNotExists}, overwriteIfExists:{overwriteIfExists}, append:{append})");
        }

        public TestFileSystem AddTestFile(string path, string content)
        {
            _allfiles.Add(path, new StringBuilder(content));
            return this;
        }
        private sealed class StringStream : Stream
        {
            // Usage:
            //var s = new StringStream(buffer);
            //s.Write(Encoding.UTF8.GetBytes("Hello Custom Stream!"));
            //Console.WriteLine(s.Result);

            private readonly StringBuilder _builder;

            public StringStream(StringBuilder builder)
            {
                _builder = builder;
            }

            //public string Result => _builder.ToString();

            public override bool CanWrite => true;
            public override bool CanRead => false;
            public override bool CanSeek => false;

            public override void Write(byte[] buffer, int offset, int count)
            {
                var text = Encoding.UTF8.GetString(buffer, offset, count);
                _builder.Append(text);
            }

            public override void Flush() { }

            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }
    }
}