' Parity script: the expected output (ComParity_ScriptingFileSystemObject.expected.txt) was captured with Windows cscript.exe and the real COM components.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Out TypeName(fso) & "|" & VarType(fso)
Out fso.BuildPath("C:\a", "b.txt") & "|" & fso.BuildPath("C:\a\", "b.txt") & "|" & fso.BuildPath("C:\a", "\b.txt")
Out fso.GetFileName("C:\a\b.txt") & "|" & fso.GetBaseName("C:\a\b.txt") & "|" & fso.GetExtensionName("C:\a\b.txt") & "|" & fso.GetParentFolderName("C:\a\b.txt")
Out fso.GetFileName("C:\a\dir\") & "|" & fso.GetExtensionName("C:\a\noext") & "|" & fso.GetParentFolderName("C:\")
Out Len(fso.GetTempName()) & "|" & Right(fso.GetTempName(), 4) & "|" & Left(fso.GetTempName(), 3)
Dim tmp : tmp = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton_parity_" & fso.GetTempName())
Dim ts : Set ts = fso.CreateTextFile(tmp)
Out TypeName(ts)
ts.WriteLine "line1" : ts.Write "line2" : ts.Close
Out fso.FileExists(tmp)
Set ts = fso.OpenTextFile(tmp, 1)
Out ts.ReadLine() & "|" & ts.AtEndOfStream & "|" & ts.ReadAll() & "|" & ts.AtEndOfStream
ts.Close
Dim f : Set f = fso.GetFile(tmp)
Out TypeName(f) & "|" & (f.Name = fso.GetFileName(tmp))
On Error Resume Next
fso.CreateTextFile(tmp) : Out "exists:" & Err.Number : Err.Clear
fso.DeleteFile tmp
fso.DeleteFile tmp : Out "delete missing:" & Err.Number : Err.Clear
fso.GetFile tmp : Out "getfile missing:" & Err.Number : Err.Clear
On Error GoTo 0
Out fso.FileExists(tmp) & "|" & TypeName(fso.GetSpecialFolder(2))
On Error Resume Next
fso.OpenTextFile tmp, 1 : Out "open missing:" & Err.Number : Err.Clear
fso.GetFolder tmp : Out "getfolder missing:" & Err.Number : Err.Clear
fso.DeleteFolder tmp : Out "deletefolder missing:" & Err.Number : Err.Clear
Set ts = fso.CreateTextFile(tmp) : ts.Close
fso.CreateTextFile tmp, False : Out "no overwrite:" & Err.Number : Err.Clear
fso.DeleteFile tmp
On Error GoTo 0
Out TypeName(fso.GetFolder(fso.GetSpecialFolder(2))) & "|" & VarType(fso.GetSpecialFolder(2))
Out fso.GetParentFolderName("C:\a") & "|" & fso.GetDriveName("C:\a\b") & "|" & fso.GetBaseName("x.tar.gz") & "|" & fso.GetFileName("a/b/c.txt") & "|" & fso.GetParentFolderName("rel")
