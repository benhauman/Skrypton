using System;
using System.Collections.Generic;

namespace Skrypton.Tests.RuntimeSupport.Components;

public interface IHostFileSystemHostService
{
    bool FileExists(string path);
    bool DirectoryExists(string path);
    System.IO.StreamReader OpenTextFileRead(string path);
    System.IO.Stream OpenTextFileWrite(string path, bool createIfNotExists, bool overwriteIfExists, bool append);
    HostFileSystemDirectoryInfo CreateDirectory(string path);
    void DeleteFile(string path);
    void DeleteDirectory(string path, bool recursive);
    void MoveFile(string src, string dst);
    void MoveDirectory(string src, string dst);
    void CopyFile(string src, string dst, bool overwrite);

    IEnumerable<HostFileSystemFileInfo> GetFiles(string directory);
    IEnumerable<HostFileSystemDirectoryInfo> GetDirectories(string directory);
    bool DriveExists(string path); // OS‑dependent implementation

    HostFileSystemDirectoryInfo GetDirectoryInfo(string path);
    HostFileSystemFileInfo GetFileInfo(string path);
    void CopyDirectory(string sourcePath, string newPath, bool overwrite);
}
public sealed class HostFileSystemDirectoryInfo
{
    public string Path { get; }
    public string Name { get; }
    public bool Exists { get; }

    public HostFileSystemDirectoryInfo(string path, string name, bool exists)
    {
        Path = path ?? throw new ArgumentNullException(nameof(path));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Exists = exists;
    }
}
public sealed class HostFileSystemFileInfo
{
    public string Path { get; }
    public string Name { get; }
    public bool Exists { get; }

    public HostFileSystemFileInfo(string path, string name, bool exists)
    {
        Path = path ?? throw new ArgumentNullException(nameof(path));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Exists = exists;
    }
}
public abstract class HostFileSystemHostServiceBase
{
    protected HostFileSystemHostServiceBase()
    {
    }
}