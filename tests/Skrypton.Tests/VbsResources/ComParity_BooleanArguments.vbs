' Parity script: the expected output (ComParity_BooleanArguments.expected.txt) was captured with Windows cscript.exe.
' The test runs this script through Skrypton (Out writes a line) and compares the output line by line.
' A number or a string passed for a Boolean parameter of a built-in is converted as COM converts it (issue #138)
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim p : p = fso.BuildPath(fso.GetSpecialFolder(2), "skrypton-bool-probe.txt")
Dim f
Set f = fso.OpenTextFile(p, 2, 1) : f.Write "a" : f.Close
Set f = fso.OpenTextFile(p, 1, 0) : Out "OpenTextFile(p, 1, 0): " & f.ReadAll : f.Close
Set f = fso.OpenTextFile(p, 8, -1) : f.Write "b" : f.Close
Set f = fso.OpenTextFile(p, 1, "False") : Out "OpenTextFile(p, 1, ""False""): " & f.ReadAll : f.Close
Set f = fso.OpenTextFile(p, 1, "true", 0) : Out "OpenTextFile(p, 1, ""true"", 0): " & f.ReadAll : f.Close
Set f = fso.CreateTextFile(p, 1) : f.Write "c" : f.Close
Out "CreateTextFile(p, 1): " & fso.OpenTextFile(p).ReadAll
Set f = fso.CreateTextFile(p, -1, 1) : f.Write "u" : f.Close
Out "CreateTextFile(p, -1, 1): " & fso.GetFile(p).Size
Set f = fso.CreateTextFile(p, "True", "0") : f.Write "d" : f.Close
Out "CreateTextFile(p, ""True"", ""0""): " & fso.GetFile(p).Size
fso.CopyFile p, p & ".copy", 1
Out "CopyFile 1: " & fso.FileExists(p & ".copy")
fso.DeleteFile p & ".copy", 0
fso.DeleteFile p, 0
Out "DeleteFile(p, 0): " & fso.FileExists(p)
Dim d : Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1 : d.Add "A", 1
Out "Dictionary CompareMode 1: " & d.Exists("a")
On Error Resume Next
Set f = fso.OpenTextFile(p, 2, "yes")
Out "OpenTextFile(p, 2, ""yes""): " & Err.Number & " " & Err.Description : Err.Clear
Set f = fso.OpenTextFile(p, 2, d)
Out "OpenTextFile(p, 2, object): " & Err.Number & " " & Err.Description : Err.Clear
Set f = fso.OpenTextFile(p, 2, Null)
Out "OpenTextFile(p, 2, Null): " & Err.Number & " " & Err.Description : Err.Clear
Out "exists after errors: " & fso.FileExists(p)
On Error GoTo 0
Dim doc : Set doc = CreateObject("Msxml2.DOMDocument.6.0")
doc.async = 0
Out "DOMDocument async 0: " & doc.async
doc.async = "True"
Out "DOMDocument async ""True"": " & doc.async
doc.preserveWhiteSpace = 1
Out "preserveWhiteSpace 1: " & doc.preserveWhiteSpace & " " & doc.loadXML("<a> <b/> </a>")
Dim d2 : Set d2 = CreateObject("Scripting.Dictionary")
d2.CompareMode = "1" : d2.Add "A", 1
Out "Dictionary CompareMode ""1"": " & d2.Exists("a")
