using System;
using System.Collections.Generic;

namespace Skrypton.RuntimeSupport.Components;

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
    HostFileSystemDirectoryInfo GetSpecialFolderTemp();
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

    /// <summary>
    /// A file with its size and times (File.Size, DateCreated, DateLastModified, DateLastAccessed)
    /// </summary>
    public HostFileSystemFileInfo(string path, string name, bool exists, long size, DateTime dateCreated, DateTime dateLastModified, DateTime dateLastAccessed)
        : this(path, name, exists)
    {
        Size = size;
        DateCreated = dateCreated;
        DateLastModified = dateLastModified;
        DateLastAccessed = dateLastAccessed;
    }

    /// <summary>
    /// The size in bytes (0 if the host service does not provide it)
    /// </summary>
    public long Size { get; }

    public DateTime DateCreated { get; }

    public DateTime DateLastModified { get; }

    public DateTime DateLastAccessed { get; }
}
public abstract class HostFileSystemHostServiceBase
{
    protected HostFileSystemHostServiceBase()
    {
    }
}