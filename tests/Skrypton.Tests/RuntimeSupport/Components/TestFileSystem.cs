using System;
using Skrypton.RuntimeSupport.Components;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Skrypton.Tests.RuntimeSupport.Components;

namespace Skrypton.Tests.RuntimeSupport.Components.FileSystemSupport
{
    public sealed class TestFileSystem : IHostFileSystemHostService
    {
        private readonly Dictionary<string, StringBuilder> _allfiles = new Dictionary<string, StringBuilder>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> _folders = new Dictionary<string, HashSet<string>>();

        public TestFileSystem()
        {
        }

        void IHostFileSystemHostService.CopyDirectory(string sourcePath, string newPath, bool overwrite)
        {
            throw new NotImplementedException();
        }

        public HostFileSystemDirectoryInfo GetSpecialFolderTemp()
        {
            return new HostFileSystemDirectoryInfo("tst", "tmp", true);
        }

        void IHostFileSystemHostService.CopyFile(string src, string dst, bool overwrite)
        {
            throw new NotImplementedException();
        }

        HostFileSystemDirectoryInfo IHostFileSystemHostService.CreateDirectory(string path)
        {
            _folders.Add(path, new HashSet<string>());
            return new HostFileSystemDirectoryInfo(path, Path.GetDirectoryName(path), exists: true);
        }

        void IHostFileSystemHostService.DeleteDirectory(string path, bool recursive)
        {
            throw new NotImplementedException();
        }

        void IHostFileSystemHostService.DeleteFile(string path)
        {
            if (!_allfiles.Remove(path))
            {
                throw new FileNotFoundException("Could not be found.", fileName: path);
            }
        }

        bool IHostFileSystemHostService.DirectoryExists(string path)
        {
            if (_allfiles.Any(x => x.Key.StartsWith(path, StringComparison.OrdinalIgnoreCase)))
                return true;
            if (_folders.TryGetValue(path, out var folder))
                return true;
            return false;
        }

        bool IHostFileSystemHostService.DriveExists(string path)
        {
            throw new NotImplementedException();
        }

        bool IHostFileSystemHostService.FileExists(string path)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));
            return _allfiles.TryGetValue(path, out _);
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
            if (_allfiles.TryGetValue(path, out _))
            {
                string name = Path.GetFileName(path);
                return new HostFileSystemFileInfo(path, name, exists: true);
            }
            throw new FileNotFoundException("Could not be found.", fileName: path);
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
            // [FS].OpenTextFileWrite(path:'tst\HtmlPreview.html', createIfNotExists:True, overwriteIfExists:True, append:False)
            if (createIfNotExists)
            {
                content = new StringBuilder();
                _allfiles[path] = content;
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