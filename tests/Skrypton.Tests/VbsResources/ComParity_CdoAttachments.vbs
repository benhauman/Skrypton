' Parity script: the expected output (ComParity_CdoAttachments.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' CDO.Message.AddAttachment and Attachments (issue #175): the body parts (TypeName, FileName, ContentMediaType, ContentTransferEncoding), the
' collection (Count, Item, For Each, Delete, DeleteAll), a file: URL, the errors for a missing file and an invalid index
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim dir : dir = fso.GetSpecialFolder(2)
Dim ext, p, f
For Each ext In Array("txt", "xml", "pdf", "csv", "zip", "htm", "png", "skryptonxyz")
    p = fso.BuildPath(dir, "sk175." & ext) : Set f = fso.CreateTextFile(p, True) : f.Write "x" : f.Close
Next
Dim m : Set m = CreateObject("CDO.Message")
Dim part
For Each ext In Array("txt", "xml", "pdf", "csv", "zip", "htm", "png", "skryptonxyz")
    Set part = m.AddAttachment(fso.BuildPath(dir, "sk175." & ext))
    Out ext & ": " & TypeName(part) & " [" & part.FileName & "] [" & part.ContentMediaType & "] [" & part.ContentTransferEncoding & "]"
Next
Out TypeName(m.Attachments) & " " & m.Attachments.Count & " " & m.Attachments(1).FileName & " " & m.Attachments.Item(2).FileName
Dim a, names : For Each a In m.Attachments : names = names & a.FileName & "," : Next : Out names
m.Attachments.Delete 1
Out m.Attachments.Count & " " & m.Attachments(1).FileName
Set part = m.AddAttachment("file:///" & Replace(fso.BuildPath(dir, "sk175.txt"), "\", "/"))
Out "file url: " & part.FileName & " " & m.Attachments.Count
m.Attachments.DeleteAll
Out "after DeleteAll: " & m.Attachments.Count
On Error Resume Next
m.AddAttachment fso.BuildPath(dir, "sk175-missing.txt")
Out "missing: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "] [" & Err.Source & "]" : Err.Clear
m.Attachments.Delete 5
Out "delete 5: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
Dim z : Set z = m.Attachments(1)
Out "item 1 of none: " & Hex(Err.Number) & " [" & Replace(Replace(Err.Description, vbCr, "\r"), vbLf, "\n") & "]" : Err.Clear
For Each ext In Array("txt", "xml", "pdf", "csv", "zip", "htm", "png", "skryptonxyz") : fso.DeleteFile fso.BuildPath(dir, "sk175." & ext) : Next
