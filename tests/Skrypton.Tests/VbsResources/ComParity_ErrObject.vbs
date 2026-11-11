' Parity script: the expected output (ComParity_ErrObject.expected.txt) was captured with Windows cscript.exe.
' Err.Number, Err.Description and Err.Source of trapped errors (issue #85) - the description of error 424 names nothing here (issue #84).
Dim x, fso, o, h, s
Sub Report(caseName)
    Out caseName & ": " & Err.Number & "|" & Err.Description & "|" & Err.Source
    Err.Clear
End Sub
Sub LocalNothing()
    Dim l
    Set l = Nothing
    On Error Resume Next
    l.Go : Report "local Nothing call"
    x = l.Name : Report "local Nothing get"
    l.Name = 1 : Report "local Nothing set"
End Sub
Function GetNothing()
    Set GetNothing = Nothing
End Function
On Error Resume Next
x = 1 / 0 : Report "division by zero"
x = CInt("abc") : Report "type mismatch"
x = CInt(99999) : Report "overflow"
Err.Raise 5 : Report "raise 5"
Err.Raise 1000, "MySource", "My description" : Report "raise with source and description"
Err.Raise vbObjectError + 1, , "custom" : Report "raise vbObjectError without source"
Err.Raise 70 : Report "raise 70"
Set fso = CreateObject("Scripting.FileSystemObject")
fso.GetFile "C:\no\such\folder\file.txt" : Report "file not found"
fso.GetFolder "C:\no\such\folder" : Report "path not found"
Set o = Nothing
o.Go : Report "global Nothing call"
o.Name = 1 : Report "global Nothing set"
x = h.Name : Report "Empty member"
s = "abc"
x = s.Length : Report "string member"
x = GetNothing().X : Report "call result member"
x = Array(1)(5) : Report "array index out of range"
LocalNothing
Err.Clear
Out "cleared: " & Err.Number & "|" & Err.Description & "|" & Err.Source
